const { Buffer } =      require('buffer');

const OleCompoundDoc =  require('./ole-compound-doc');
const Document =        require('./document');

const BufferReader =    require('./buffer-reader');
const FileReader =      require('./file-reader');

class WordExtractor {

  constructor() {}

  extract(source) {
    let reader = null;
    if (Buffer.isBuffer(source)) {
      reader = new BufferReader(source);
    } else if (typeof source === 'string') {
      reader = new FileReader(source);
    }
    const document = new OleCompoundDoc(reader);
    return document.read()
      .then(() =>
        this.documentStream(document, 'WordDocument')
          .then((stream) => this.streamBuffer(stream))
          .then((buffer) => this.extractWordDocument(document, buffer))
      )
      .finally(() => reader.close());
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

  writeBookmarks(buffer, tableBuffer, result) {
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
      result.bookmarks[segment] = {start: cpStart, end: cpEnd};
      offset = offset + length + 2;
    }
  }

  writePieces(buffer, tableBuffer, result) {
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
      result.pieces.push(piece);

      start = start + (unicode ? Math.floor(totLength / 2) : totLength);
      lastPosition = lastPosition + piece.length;
    }
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
        const result = new Document();
        result.boundaries.fcMin = buffer.readUInt32LE(0x0018);
        result.boundaries.ccpText = buffer.readUInt32LE(0x004c);
        result.boundaries.ccpFtn = buffer.readUInt32LE(0x0050);
        result.boundaries.ccpHdd = buffer.readUInt32LE(0x0054);
        result.boundaries.ccpAtn = buffer.readUInt32LE(0x005c);
        result.boundaries.ccpEdn = buffer.readUInt32LE(0x0060);

        this.writeBookmarks(buffer, streamBuffer, result);
        this.writePieces(buffer, streamBuffer, result);

        return result;
      });
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
      return piece.text = this.addText(buffer, textStart, textEnd);
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


module.exports = WordExtractor;
