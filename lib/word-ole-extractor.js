/**
 * @module word-ole-extractor
 * 
 * @description
 * Implements the main logic of extracting text from "classic" OLE-based Word files.
 * Depends on [OleCompoundDoc]{@link module:ole-compound-doc~OleCompoundDoc} 
 * for most of the underlying OLE logic. Note that
 * [OpenOfficeExtractor]{@link module:open-office-extractor~OpenOfficeExtractor} is 
 * used for newer, Open Office-style, files. 
 */

const OleCompoundDoc = require('./ole-compound-doc');
const Document = require('./document');
const { replace, binaryToUnicode, clean } = require('./filters');

const sprmCFRMarkDel = 0x00;

// /**
//  * Given a cp-style file offset, finds the containing piece index.
//  * @param {*} offset the character offset
//  * @returns the piece index
//  */
const getPieceIndex = (pieces, position) => {
  for (let i = 0; i < pieces.length; i++) {
    const piece = pieces[i];
    if (position <= piece.endCp) { 
      return i; 
    }
  }
};

// /**
//  * Given a file-style offset, finds the containing piece index.
//  * @param {*} offset the character offset
//  * @returns the piece index
//  */
const getPieceIndexByFilePos = (pieces, position) => {
  for (let i = 0; i < pieces.length; i++) {
    const piece = pieces[i];
    if (position <= piece.endFilePos) { 
      return i; 
    }
  }
};

// /**
//  * Given a stream-style file offset, finds the containing piece index.
//  * @param {*} offset the character offset
//  * @returns the piece index
//  */
const getPieceStream = (pieces, position) => {    // eslint-disable-line no-unused-vars
  for (let i = 0; i < pieces.length; i++) {
    const piece = pieces[i];
    if (position <= piece.endStream) { 
      return i; 
    }
  }
};

const processSprms = (buffer, offset, handler) => {
  while (offset < buffer.length - 1) {
    const sprm = buffer.readUInt16LE(offset);
    const ispmd = sprm & 0x1f;
    const fspec = (sprm >> 9) & 0x01;
    const sgc = (sprm >> 10) & 0x07;
    const spra = (sprm >> 13) & 0x07;

    offset += 2;

    handler(buffer, offset, sprm, ispmd, fspec, sgc, spra);

    if (spra === 0) {
      offset += 1;
      continue;
    } else if (spra === 1) {
      offset += 1;
      continue;
    } else if (spra === 2) {
      offset += 2;
      continue;
    } else if (spra === 3) {
      offset += 4;
      continue;
    } else if (spra === 4 || spra === 5) {
      offset += 2;
      continue;
    } else if (spra === 6) {
      offset += buffer.readUInt8(offset) + 1;
      continue;
    } else if (spra === 7) {
      offset += 3;
      continue;
    } else {
      throw new Error("Unparsed sprm");
    }

  }
};

/**
 * @class
 * The main class implementing extraction from OLE-based Word files.
 */
class WordOleExtractor {

  constructor() { 
    this._pieces = [];
    this._bookmarks = {};
    this._boundaries = {};
  }

  /**
   * The main extraction method. This creates an OLE compound document
   * interface, then opens up a stream and extracts out the main
   * stream.
   * @param {*} reader 
   * @returns 
   */
  extract(reader) {
    const document = new OleCompoundDoc(reader);
    return document.read()
      .then(() =>
        this.documentStream(document, 'WordDocument')
          .then((stream) => this.streamBuffer(stream))
          .then((buffer) => this.extractWordDocument(document, buffer))
      );
  }

  /**
   * Reads and extracts a character range from the pieces.
   * @param {*} start the start offset
   * @param {*} end the end offset
   * @returns a character string
   */
  getTextRange(start, end) {
    const pieces = this._pieces;
    const startPiece = getPieceIndex(pieces, start);
    const endPiece = getPieceIndex(pieces, end);
    const result = [];
    for (let i = startPiece, end1 = endPiece; i <= end1; i++) {
      const piece = pieces[i];
      const xstart = i === startPiece ? start - piece.startCp : 0;
      const xend = i === endPiece ? end - piece.startCp : piece.endCp;
      result.push(replace(piece.text.substring(xstart, xend)));
    }

    return clean(result.join(""));
  }

  replaceSelectedRange(start, end, character) {
    const pieces = this._pieces;
    const startPiece = getPieceIndex(pieces, start);
    const endPiece = getPieceIndex(pieces, end);
    for (let i = startPiece, end1 = endPiece; i <= end1; i++) {
      const piece = pieces[i];
      this.fillPieceRange(piece, start, end, character);
    }
  }

  replaceSelectedRangeByFilePos(start, end, character) {
    const pieces = this._pieces;
    const startPiece = getPieceIndexByFilePos(pieces, start);
    const endPiece = getPieceIndexByFilePos(pieces, end);
    for (let i = startPiece, end1 = endPiece; i <= end1; i++) {
      const piece = pieces[i];
      this.fillPieceRangeByFilePos(piece, start, end, character);
    }
  }

  markDeletedRange(start, end) {
    this.replaceSelectedRangeByFilePos(start, end, '\x00');
  }

  /**
   * Builds and returns a {@link Document} object corresponding to the text
   * in the original document. This involves reading and retrieving the text
   * ranges corresponding to the primary document parts. The text segments are
   * read from the extracted table of text pieces.
   * @returns a {@link Document} object
   */
  buildDocument() {
    const document = new Document();
    let start = 0;

    document.setBody(this.getTextRange(start, start + this._boundaries.ccpText));
    start += this._boundaries.ccpText;

    document.setFootnotes(this.getTextRange(start, start + this._boundaries.ccpFtn));
    start += this._boundaries.ccpFtn;

    document.setHeaders(this.getTextRange(start, start + this._boundaries.ccpHdd));
    start += this._boundaries.ccpHdd;

    document.setAnnotations(this.getTextRange(start, start + this._boundaries.ccpAtn));
    start += this._boundaries.ccpAtn;

    document.setEndnotes(this.getTextRange(start, start + this._boundaries.ccpEdn));
    start += this._boundaries.ccpEdn;

    return document;
  }


  /**
   * Main logic top level function for unpacking a Word document 
   * @param {*} document the OLE document
   * @param {*} buffer a buffer 
   * @returns a Promise which resolves to a {@link Document}
   */
  extractWordDocument(document, buffer) {
    const magic = buffer.readUInt16LE(0);
    if (magic !== 0xa5ec) {
      return Promise.reject(new Error(`This does not seem to be a Word document: Invalid magic number: ${magic.toString(16)}`));
    }

    const flags = buffer.readUInt16LE(0xA);

    const streamName = (flags & 0x0200) !== 0 ? "1Table" : "0Table";

    return this.documentStream(document, streamName)
      .then((stream) => this.streamBuffer(stream))
      .then((streamBuffer) => {
        this._boundaries.fcMin = buffer.readUInt32LE(0x0018);
        this._boundaries.ccpText = buffer.readUInt32LE(0x004c);
        this._boundaries.ccpFtn = buffer.readUInt32LE(0x0050);
        this._boundaries.ccpHdd = buffer.readUInt32LE(0x0054);
        this._boundaries.ccpAtn = buffer.readUInt32LE(0x005c);
        this._boundaries.ccpEdn = buffer.readUInt32LE(0x0060);

        this.writeBookmarks(buffer, streamBuffer);
        this.writePieces(buffer, streamBuffer);
        this.writeCharacterProperties(buffer, streamBuffer);
        this.writeParagraphProperties(buffer, streamBuffer);
        //this.writeFields(buffer, streamBuffer, result);

        return this.buildDocument();
      });
  }

  /**
   * Given a piece, and a starting and ending cp-style file offset, 
   * and a replacement character, updates the piece text to replace
   * between start and end with the given character.
   * @param {*} piece the piece
   * @param {*} start the starting character offset
   * @param {*} end the endingcharacter offset
   * @param {*} character the replacement character
   */
  fillPieceRange(piece, start, end, character) {
    const pieceStart = piece.startCp;
    const pieceEnd = pieceStart + piece.length;
    const original = piece.text;
    if (start < pieceStart) start = pieceStart;
    if (end > pieceEnd) end = pieceEnd;
    const modified = 
      ((start == pieceStart) ? '' : original.slice(0, start - pieceStart)) +
      ''.padStart(end - start, character) +
      ((end == pieceEnd) ? '' : original.slice(end - pieceEnd));
    piece.text = modified;
  }

  /**
   * Given a piece, and a starting and ending filePos-style file offset, 
   * and a replacement character, updates the piece text to replace
   * between start and end with the given character. This is used when
   * applying character styles, which use filePos values rather than cp
   * values.
   *
   * @param {*} piece the piece
   * @param {*} start the starting character offset
   * @param {*} end the endingcharacter offset
   * @param {*} character the replacement character
   */
  fillPieceRangeByFilePos(piece, start, end, character) {
    const pieceStart = piece.startFilePos;
    const pieceEnd = pieceStart + piece.size;
    const original = piece.text;
    if (start < pieceStart) start = pieceStart;
    if (end > pieceEnd) end = pieceEnd;
    const modified = 
      ((start == pieceStart) ? '' : original.slice(0, (start - pieceStart) / piece.bpc)) +
      ''.padStart((end - start) / piece.bpc, character) +
      ((end == pieceEnd) ? '' : original.slice((end - pieceEnd) / piece.bpc));
    piece.text = modified;
  }

  documentStream(document, streamName) {
    return Promise.resolve(document.stream(streamName));
  }

  streamBuffer(stream) {
    return new Promise((resolve, reject) => {
      const chunks = [];
      stream.on('data', (chunk) => chunks.push(chunk));
      stream.on('error', (error) => reject(error));
      stream.on('end', () => resolve(Buffer.concat(chunks)));
      return stream;
    });
  }

  writeFields(buffer, tableBuffer, result) {    // eslint-disable-line no-unused-vars
    const fcPlcffldMom = buffer.readInt32LE(0x011a);
    const lcbPlcffldMom = buffer.readUInt32LE(0x011e);
    //console.log(fcPlcffldMom, lcbPlcffldMom, tableBuffer.length);

    if (lcbPlcffldMom == 0) {
      return;
    }

    const fieldCount = (lcbPlcffldMom - 4) / 6;
    //console.log("extracting", fieldCount, "fields");

    const dataOffset = (fieldCount + 1) * 4;

    const plcffldMom = tableBuffer.slice(fcPlcffldMom, fcPlcffldMom + lcbPlcffldMom);
    for(let i = 0; i < fieldCount; i++) {
      const cp = plcffldMom.readUInt32LE(i * 4);      // eslint-disable-line no-unused-vars
      const fld = plcffldMom.readUInt16LE(dataOffset + i * 2);
      const byte1 = fld & 0xff;
      const byte2 = fld >> 8;                         // eslint-disable-line no-unused-vars
      if ((byte1 & 0x1f) == 19) {
        //console.log("A", i, cp, byte1.toString(16), byte2.toString(16));
      } else {
        //console.log("B", i, cp, byte1.toString(16), byte2.toString(16));
      }
    }
  }

  /**
   * Extracts and stores the document bookmarks into a local field.
   * @param {*} buffer 
   * @param {*} tableBuffer 
   */
  writeBookmarks(buffer, tableBuffer) {
    const fcSttbfBkmk = buffer.readUInt32LE(0x0142);
    const lcbSttbfBkmk = buffer.readUInt32LE(0x0146);
    const fcPlcfBkf = buffer.readUInt32LE(0x014a);
    const lcbPlcfBkf = buffer.readUInt32LE(0x014e);
    const fcPlcfBkl = buffer.readUInt32LE(0x0152);
    const lcbPlcfBkl = buffer.readUInt32LE(0x0156);

    if (lcbSttbfBkmk === 0) { 
      return; 
    }

    const sttbfBkmk = tableBuffer.slice(fcSttbfBkmk, fcSttbfBkmk + lcbSttbfBkmk);
    const plcfBkf = tableBuffer.slice(fcPlcfBkf, fcPlcfBkf + lcbPlcfBkf);
    const plcfBkl = tableBuffer.slice(fcPlcfBkl, fcPlcfBkl + lcbPlcfBkl);

    const fcExtend = sttbfBkmk.readUInt16LE(0);
    const cData = sttbfBkmk.readUInt16LE(2);       // eslint-disable-line no-unused-vars
    const cbExtra = sttbfBkmk.readUInt16LE(4);     // eslint-disable-line no-unused-vars

    if (fcExtend !== 0xffff) {
      throw new Error("Internal error: unexpected single-byte bookmark data");
    }

    let offset = 6;
    const index = 0;

    while (offset < lcbSttbfBkmk) {
      let length = sttbfBkmk.readUInt16LE(offset);
      length = length * 2;
      const segment = sttbfBkmk.slice(offset + 2, offset + 2 + length);
      const cpStart = plcfBkf.readUInt32LE(index * 4);
      const cpEnd = plcfBkl.readUInt32LE(index * 4);
      this._bookmarks[segment] = {start: cpStart, end: cpEnd};
      offset = offset + length + 2;
    }
  }

  /**
   * Extracts and stores the document text pieces into a local field. This is
   * probably the most crucial part of text extraction, as it is where we
   * get text corresponding to character positions. These may be stored in a 
   * different order in the file compared to the order we want them. 
   * 
   * @param {*} buffer 
   * @param {*} tableBuffer 
   */
  writePieces(buffer, tableBuffer) {
    let flag;
    let pos = buffer.readUInt32LE(0x01a2);

    while (true) {                          // eslint-disable-line no-constant-condition
      flag = tableBuffer.readUInt8(pos);
      if (flag !== 1) { break; }

      pos = pos + 1;
      const skip = tableBuffer.readUInt16LE(pos);
      pos = pos + 2 + skip;
    }

    flag = tableBuffer.readUInt8(pos);
    pos = pos + 1;
    if (flag !== 2) {
      throw new Error("Internal error: ccorrupted Word file");
    }

    const pieceTableSize = tableBuffer.readUInt32LE(pos);
    pos = pos + 4;

    const pieces = (pieceTableSize - 4) / 12;

    let startCp = 0;
    let startStream = 0;

    for (let x = 0, end = pieces - 1; x <= end; x++) {
      const offset = pos + ((pieces + 1) * 4) + (x * 8) + 2;
      let startFilePos = tableBuffer.readUInt32LE(offset);
      let unicode = false;
      if ((startFilePos & 0x40000000) === 0) {
        unicode = true;
      } else {
        startFilePos = startFilePos & ~(0x40000000);
        startFilePos = Math.floor(startFilePos / 2);
      }
      const lStart = tableBuffer.readUInt32LE(pos + (x * 4));
      const lEnd = tableBuffer.readUInt32LE(pos + ((x + 1) * 4));
      const totLength = lEnd - lStart;

      const piece = {
        startCp,
        startStream,
        totLength,
        startFilePos,
        unicode,
        bpc: (unicode) ? 2 : 1
      };

      if (unicode) {
        const textBuffer = buffer.slice(startFilePos, startFilePos + 2 * (lEnd - lStart));
        piece.text = textBuffer.toString('ucs2');
      } else {
        const textBuffer = buffer.slice(startFilePos, startFilePos + (lEnd - lStart));
        piece.text = binaryToUnicode(textBuffer.toString('binary'));
      }
      
      piece.length = piece.text.length;
      piece.size = piece.bpc * (lEnd - lStart); 

      piece.endCp = piece.startCp + piece.length;
      piece.endStream = piece.startStream + piece.bpc * (lEnd - lStart);
      piece.endFilePos = piece.startFilePos + piece.bpc * (lEnd - lStart);

      startCp = piece.endCp;
      startStream = piece.endStream;

      this._pieces.push(piece);
    }
  }

  writeParagraphProperties(buffer, tableBuffer) {
    const fcPlcfbtePapx = buffer.readUInt32LE(0x0102);
    const lcbPlcfbtePapx = buffer.readUInt32LE(0x0106);

    const plcBtePapxCount = (lcbPlcfbtePapx - 4) / 8;
    const dataOffset = (plcBtePapxCount + 1) * 4;
    const plcBtePapx = tableBuffer.slice(fcPlcfbtePapx, fcPlcfbtePapx + lcbPlcfbtePapx);

    for(let i = 0; i < plcBtePapxCount; i++) {
      const cp = plcBtePapx.readUInt32LE(i * 4);    // eslint-disable-line no-unused-vars
      const papxFkpBlock = plcBtePapx.readUInt32LE(dataOffset + i * 4);
      //console.log("paragraph property", cp, papxFkpBlock);

      const papxFkpBlockBuffer = buffer.slice(papxFkpBlock * 512, (papxFkpBlock + 1) * 512);
      //console.log("papxFkpBlockBuffer", papxFkpBlockBuffer);

      const crun = papxFkpBlockBuffer.readUInt8(511);
      //console.log("crun", crun);

      for(let j = 0; j < crun; j++) {
        const rgfc = papxFkpBlockBuffer.readUInt32LE(j * 4);
        const rgfcNext = papxFkpBlockBuffer.readUInt32LE((j + 1) * 4);

        const cbLocation = (crun + 1) * 4 + j * 13;
        const cbIndex = papxFkpBlockBuffer.readUInt8(cbLocation) * 2;

        const cb = papxFkpBlockBuffer.readUInt8(cbIndex);
        let grpPrlAndIstd = null;
        if (cb !== 0) {
          grpPrlAndIstd = papxFkpBlockBuffer.slice(cbIndex + 1, cbIndex + 1 + (2 * cb) - 1);
        } else {
          const cb2 = papxFkpBlockBuffer.readUInt8(cbIndex + 1);
          grpPrlAndIstd = papxFkpBlockBuffer.slice(cbIndex + 2, cbIndex + 2 + (2 * cb2));
        }
        //console.log("para; ", j, "rgfc=", rgfc, "rgfcNext=", rgfcNext, "grpPrlAndIstd=", grpPrlAndIstd);

        const istd = grpPrlAndIstd.readUInt16LE(0);    // eslint-disable-line no-unused-vars
        processSprms(grpPrlAndIstd, 2, (buffer, offset, sprm, ispmd, fspec, sgc, spra) => {    // eslint-disable-line no-unused-vars
          //console.log("sprm x", offset, sprm.toString(16), ispmd, fspec, sgc, spra);
          if (sprm === 0x2417) {
            //console.log("table row end", rgfc, rgfcNext);

            this.replaceSelectedRange(rgfc, rgfcNext, '\n');
          }
        });
      }
      
    }
  }

  writeCharacterProperties(buffer, tableBuffer) {
    const fcPlcfbteChpx = buffer.readUInt32LE(0x00fa);
    const lcbPlcfbteChpx = buffer.readUInt32LE(0x00fe);

    const plcBteChpxCount = (lcbPlcfbteChpx - 4) / 8;
    //console.log("character format runs", plcBteChpxCount, fcPlcfbteChpx, lcbPlcfbteChpx);

    const dataOffset = (plcBteChpxCount + 1) * 4;
    const plcBteChpx = tableBuffer.slice(fcPlcfbteChpx, fcPlcfbteChpx + lcbPlcfbteChpx);

    //const cpLast = plcBteChpx.readUInt32LE(plcBteChpxCount * 4);
    //console.log("last cp", cpLast);

    this._deletions = [];
    let lastDeletionEnd = null;

    for(let i = 0; i < plcBteChpxCount; i++) {
      const cp = plcBteChpx.readUInt32LE(i * 4);    // eslint-disable-line no-unused-vars
      const chpxFkpBlock = plcBteChpx.readUInt32LE(dataOffset + i * 4);
      //console.log("character property", cp, chpxFkpBlock);

      const chpxFkpBlockBuffer = buffer.slice(chpxFkpBlock * 512, (chpxFkpBlock + 1) * 512);
      //console.log("chpxFkpBlockBuffer", chpxFkpBlockBuffer);

      const crun = chpxFkpBlockBuffer.readUInt8(511);
      //console.log("crun", crun);

      for(let j = 0; j < crun; j++) {
        const rgfc = chpxFkpBlockBuffer.readUInt32LE(j * 4);
        const rgfcNext = chpxFkpBlockBuffer.readUInt32LE((j + 1) * 4);
        const rgb = chpxFkpBlockBuffer.readUInt8((crun + 1) * 4 + j);
        if (rgb == 0) {
          //console.log("skipping run; ", j, "rgfc=", rgfc, "rgb=", rgb);
          continue;
        }
        const chpxOffset = rgb * 2;
        const cb = chpxFkpBlockBuffer.readUInt8(chpxOffset);
        const grpprl = chpxFkpBlockBuffer.slice(chpxOffset + 1, chpxOffset + 1 + cb);
        //console.log("found run; ", j, "rgfc=", rgfc, "rgb=", rgb, "cb=", cb, "grpprl=", grpprl);

        processSprms(grpprl, 0, (buffer, offset, sprm, ispmd) => {
          if (ispmd === sprmCFRMarkDel) {
            //console.log("text deleted", rgfc, rgfcNext);
            const ld = this._deletions.length - 1;

            if (ld >= 0 && lastDeletionEnd === rgfc) {
              this.markDeletedRange(lastDeletionEnd, rgfcNext);
            } else {
              this.markDeletedRange(rgfc, rgfcNext);
            }
            lastDeletionEnd = rgfcNext;

            // if (ld >= 0 && this._deletions[ld].end === rgfc) {
            //   this._deletions[ld].end = rgfcNext;
            // } else {
            //   this._deletions.push({start: rgfc, end: rgfcNext});
            // }
          }
        });
      }
    }
  }

}

module.exports = WordOleExtractor;
