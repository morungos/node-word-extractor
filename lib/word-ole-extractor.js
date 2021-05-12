const OleCompoundDoc = require('./ole-compound-doc');
const Document = require('./document');
const { replace, binaryToUnicode, clean } = require('./filters');

const sprmCFRMarkDel = 0x00;

const getPieceIndex = (pieces, position) => {
  for (let i = 0; i < pieces.length; i++) {
    const piece = pieces[i];
    if (position <= piece.endPosition) { 
      return i; 
    }
  }
};

class WordOleExtractor {

  constructor() { 
    this._pieces = [];
    this._bookmarks = {};
    this._boundaries = {};
  }

  extract(reader) {
    const document = new OleCompoundDoc(reader);
    return document.read()
      .then(() =>
        this.documentStream(document, 'WordDocument')
          .then((stream) => this.streamBuffer(stream))
          .then((buffer) => this.extractWordDocument(document, buffer))
      );
  }

  getTextRange(start, end) {
    const pieces = this._pieces;
    const startPiece = getPieceIndex(pieces, start);
    const endPiece = getPieceIndex(pieces, end);
    const result = [];
    for (let i = startPiece, end1 = endPiece; i <= end1; i++) {
      const piece = pieces[i];
      const xstart = i === startPiece ? start - piece.position : 0;
      const xend = i === endPiece ? end - piece.position : piece.endPosition;
      result.push(replace(piece.text.substring(xstart, xend)));
    }

    return clean(result.join(""));
  }

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
        //this.writeCharacterProperties(buffer, streamBuffer, result);
        //this.writeFields(buffer, streamBuffer, result);

        //console.log(result);

        return this.buildDocument();
      });
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
    console.log(fcPlcffldMom, lcbPlcffldMom, tableBuffer.length);

    if (lcbPlcffldMom == 0) {
      return;
    }

    const fieldCount = (lcbPlcffldMom - 4) / 6;
    console.log("extracting", fieldCount, "fields");

    const dataOffset = (fieldCount + 1) * 4;

    const plcffldMom = tableBuffer.slice(fcPlcffldMom, fcPlcffldMom + lcbPlcffldMom);
    for(let i = 0; i < fieldCount; i++) {
      const cp = plcffldMom.readUInt32LE(i * 4);
      const fld = plcffldMom.readUInt16LE(dataOffset + i * 2);
      const byte1 = fld & 0xff;
      const byte2 = fld >> 8;
      if ((byte1 & 0x1f) == 19) {
        console.log("A", i, cp, byte1.toString(16), byte2.toString(16));
      } else {
        console.log("B", i, cp, byte1.toString(16), byte2.toString(16));
      }
    }
  }

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
    let start = 0;
    let lastPosition = 0;

    for (let x = 0, end = pieces - 1; x <= end; x++) {
      const offset = pos + ((pieces + 1) * 4) + (x * 8) + 2;
      let filePos = tableBuffer.readUInt32LE(offset);
      let unicode = false;
      if ((filePos & 0x40000000) === 0) {
        unicode = true;
      } else {
        filePos = filePos & ~(0x40000000);
        filePos = Math.floor(filePos / 2);
      }
      const lStart = tableBuffer.readUInt32LE(pos + (x * 4));
      const lEnd = tableBuffer.readUInt32LE(pos + ((x + 1) * 4));
      const totLength = lEnd - lStart;

      const piece = {
        start,
        totLength,
        filePos,
        unicode
      };

      this.getPiece(buffer, piece);
      piece.length = piece.text.length;
      piece.position = lastPosition;
      piece.endPosition = lastPosition + piece.length;
      this._pieces.push(piece);

      start = start + (unicode ? Math.floor(totLength / 2) : totLength);
      lastPosition = lastPosition + piece.length;
    }
  }

  writeCharacterProperties(buffer, tableBuffer, result) {
    const fcPlcfbteChpx = buffer.readUInt32LE(0x00fa);
    const lcbPlcfbteChpx = buffer.readUInt32LE(0x00fe);

    const plcBteChpxCount = (lcbPlcfbteChpx - 4) / 8;
    console.log("character format runs", plcBteChpxCount, fcPlcfbteChpx, lcbPlcfbteChpx);

    const dataOffset = (plcBteChpxCount + 1) * 4;
    const plcBteChpx = tableBuffer.slice(fcPlcfbteChpx, fcPlcfbteChpx + lcbPlcfbteChpx);

    const cpLast = plcBteChpx.readUInt32LE(plcBteChpxCount * 4);
    console.log("last cp", cpLast);

    result.deletions = [];

    for(let i = 0; i < plcBteChpxCount; i++) {
      const cp = plcBteChpx.readUInt32LE(i * 4);
      const chpxFkpBlock = plcBteChpx.readUInt32LE(dataOffset + i * 4);
      console.log("character property", cp, chpxFkpBlock);

      const chpxFkpBlockBuffer = buffer.slice(chpxFkpBlock * 512, (chpxFkpBlock + 1) * 512);
      console.log("chpxFkpBlockBuffer", chpxFkpBlockBuffer);

      const crun = chpxFkpBlockBuffer.readUInt8(511);
      console.log("crun", crun);

      for(let j = 0; j < crun; j++) {
        const rgfc = chpxFkpBlockBuffer.readUInt32LE(j * 4);
        const rgfcNext = chpxFkpBlockBuffer.readUInt32LE((j + 1) * 4);
        const rgb = chpxFkpBlockBuffer.readUInt8((crun + 1) * 4 + j);
        if (rgb == 0) {
          console.log("skipping run; ", j, "rgfc=", rgfc, "rgb=", rgb);
          continue;
        }
        const chpxOffset = rgb * 2;
        const cb = chpxFkpBlockBuffer.readUInt8(chpxOffset);
        const grpprl = chpxFkpBlockBuffer.slice(chpxOffset + 1, chpxOffset + 1 + cb);
        console.log("found run; ", j, "rgfc=", rgfc, "rgb=", rgb, "cb=", cb, "grpprl=", grpprl);

        let k = 0;
        while (k < grpprl.length) {
          const sprm = grpprl.readUInt16LE(k);
          const ispmd = sprm & 0x1f;
          const fspec = (sprm >> 9) & 0x01;
          const sgc = (sprm >> 10) & 0x07;
          const spra = (sprm >> 13) & 0x07;

          console.log("sprm", k, sprm.toString(16), ispmd, fspec, sgc, spra);

          if (ispmd === sprmCFRMarkDel) {
            console.log("text deleted", rgfc, rgfcNext);
            const ld = result.deletions.length - 1;
            if (ld >= 0 && result.deletions[ld].end === rgfc) {
              result.deletions[ld].end = rgfcNext;
            } else {
              result.deletions.push({start: rgfc, end: rgfcNext});
            }
          }

          k += 2;
          if (spra === 0) {
            k += 1;
            continue;
          } else if (spra === 1) {
            k += 1;
            continue;
          } else if (spra === 2) {
            k += 2;
            continue;
          } else if (spra === 3) {
            k += 4;
            continue;
          } else if (spra === 4 || spra === 5) {
            k += 2;
            continue;
          } else if (spra === 6) {
            k += 3;
            continue;
          } else if (spra === 7) {
            k += chpxFkpBlockBuffer.readUInt8(k);
            continue;
          } else {
            throw new Error("Unparsed sprm");
          }

        }
      }
    }
  }

  getPiece(buffer, piece) {
    const pstart = piece.start;
    const ptotLength = piece.totLength;
    const pfilePos = piece.filePos;
    const punicode = piece.unicode;

    const pend = pstart + ptotLength;
    const textStart = pfilePos;
    const textEnd = textStart + (pend - pstart);

    if (punicode) {
      return piece.text = this.addUnicodeText(buffer, textStart, textEnd);
    } else {
      return piece.text = binaryToUnicode(this.addText(buffer, textStart, textEnd));
    }
  }

  addText(buffer, textStart, textEnd) {
    const slice = buffer.slice(textStart, textEnd);
    return slice.toString('binary');
  }

  addUnicodeText(buffer, textStart, textEnd) {
    const slice = buffer.slice(textStart, (2*textEnd) - textStart);
    const string = slice.toString('ucs2');

    // See the conversion table for FcCompressed structures. Note that these
    // should not affect positions, as these are characters now, not bytes
    // for i in [0..string.length]
    //   if

    return string;
  }

}

module.exports = WordOleExtractor;
