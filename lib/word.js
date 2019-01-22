/*
 * decaffeinate suggestions:
 * DS102: Remove unnecessary code created because of implicit returns
 * DS205: Consider reworking code to avoid use of IIFEs
 * DS206: Consider reworking classes to avoid initClass
 * Full docs: https://github.com/decaffeinate/decaffeinate/blob/master/docs/suggestions.md
 */
const { Buffer } =     require('buffer');

const oleDoc =         require('./ole-doc').OleCompoundDoc;

const filters =        require('./filters');         // eslint-disable-line no-unused-vars
const translations =   require('./translations');    // eslint-disable-line no-unused-vars

const Document =       require('./document');

var WordExtractor = (function() {
  let streamBuffer = undefined;
  let extractDocument = undefined;
  let documentStream = undefined;
  let writeBookmarks = undefined;
  let writePieces = undefined;
  let extractWordDocument = undefined;
  let getPiece = undefined;
  let addText = undefined;
  let addUnicodeText = undefined;
  WordExtractor = class WordExtractor {
    static initClass() {
  
  
      //# Given an OLE stream, returns all the data in a buffer,
      //# as a promise.
      streamBuffer = stream =>
        new Promise(function(resolve, reject) {
          const chunks = [];
          stream.on('data', chunk => chunks.push(chunk));
          stream.on('error', error => reject(error));
          return stream.on('end', () => resolve(Buffer.concat(chunks)));
        })
      ;
  
  
      extractDocument = filename =>
        new Promise(function(resolve, reject) {
          const document = new oleDoc(filename);
          document.on('err', error => {
            return reject(error);
          });
          document.on('ready', () => {
            return resolve(document);
          });
          return document.read();
        })
      ;
  
  
      documentStream = (document, stream) => Promise.resolve(document.stream(stream));
  
  
      writeBookmarks = function(buffer, tableBuffer, result) {
        const fcSttbfBkmk = buffer.readUInt32LE(0x0142);
        const lcbSttbfBkmk = buffer.readUInt32LE(0x0146);
        const fcPlcfBkf = buffer.readUInt32LE(0x014a);
        const lcbPlcfBkf = buffer.readUInt32LE(0x014e);
        const fcPlcfBkl = buffer.readUInt32LE(0x0152);
        const lcbPlcfBkl = buffer.readUInt32LE(0x0156);
  
        if (lcbSttbfBkmk === 0) { return; }
  
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
        const bookmarks = {};                          // eslint-disable-line no-unused-vars
  
        while (offset < lcbSttbfBkmk) {
          let length = sttbfBkmk.readUInt16LE(offset);
          length = length * 2;
          const segment = sttbfBkmk.slice(offset + 2, offset + 2 + length);
          const cpStart = plcfBkf.readUInt32LE(index * 4);
          const cpEnd = plcfBkl.readUInt32LE(index * 4);
          result.bookmarks[segment] = {start: cpStart, end: cpEnd};
          offset = offset + length + 2;
        }
      };
  
  
      writePieces = function(buffer, tableBuffer, result) {
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

          getPiece(buffer, piece);
          piece.length = piece.text.length;
          piece.position = lastPosition;
          piece.endPosition = lastPosition + piece.length;
          result.pieces.push(piece);

          start = start + (unicode ? Math.floor(totLength / 2) : totLength);
          lastPosition = lastPosition + piece.length;
        }
      };
  
      extractWordDocument = (document, buffer) =>
        new Promise(function(resolve, reject) {
          const magic = buffer.readUInt16LE(0);
          if (magic !== 0xa5ec) {
            return reject(new Error(`This does not seem to be a Word document: Invalid magic number: ${magic.toString(16)}`));
          }
  
          const flags = buffer.readUInt16LE(0xA);
  
          const table = (flags & 0x0200) !== 0 ? "1Table" : "0Table";
  
          return documentStream(document, table)
            .then(stream => streamBuffer(stream)).then(function(tableBuffer) {
              const result = new Document();
              result.boundaries.fcMin = buffer.readUInt32LE(0x0018);
              result.boundaries.ccpText = buffer.readUInt32LE(0x004c);
              result.boundaries.ccpFtn = buffer.readUInt32LE(0x0050);
              result.boundaries.ccpHdd = buffer.readUInt32LE(0x0054);
              result.boundaries.ccpAtn = buffer.readUInt32LE(0x005c);
              result.boundaries.ccpEdn = buffer.readUInt32LE(0x0060);
  
              writeBookmarks(buffer, tableBuffer, result);
              writePieces(buffer, tableBuffer, result);
  
              return resolve(result);}).catch(error => reject(error));
        })
      ;
  
  
      getPiece = function(buffer, piece) {
        const pstart = piece.start;
        const ptotLength = piece.totLength;
        const pfilePos = piece.filePos;
        const punicode = piece.unicode;
  
        const pend = pstart + ptotLength;
        const textStart = pfilePos;
        const textEnd = textStart + (pend - pstart);
  
        if (punicode) {
          return piece.text = addUnicodeText(buffer, textStart, textEnd);
        } else {
          return piece.text = addText(buffer, textStart, textEnd);
        }
      };
  
  
      addText = function(buffer, textStart, textEnd) {
        const slice = buffer.slice(textStart, textEnd);
        return slice.toString('binary');
      };
  
      addUnicodeText = function(buffer, textStart, textEnd) {
        const slice = buffer.slice(textStart, (2*textEnd) - textStart);
        const string = slice.toString('ucs2');
  
        // See the conversion table for FcCompressed structures. Note that these
        // should not affect positions, as these are characters now, not bytes
        // for i in [0..string.length]
        //   if
  
        return string;
      };
    }

    constructor() {}


    extract(filename) {
      return extractDocument(filename)
        .then(document =>
          documentStream(document, 'WordDocument')
            .then(stream => streamBuffer(stream)).then(buffer => extractWordDocument(document, buffer))
        );
    }
  };
  WordExtractor.initClass();
  return WordExtractor;
})();


module.exports = WordExtractor;
