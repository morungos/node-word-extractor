var Buffer, Promise, Word, filters, oleDoc, translations,
  slice1 = [].slice;

Buffer = require('buffer').Buffer;

oleDoc = require('ole-doc').OleCompoundDoc;

Promise = require('bluebird');

filters = [];

filters[0x0002] = 0;

filters[0x0005] = 0;

filters[0x0008] = 0;

filters[0x2018] = "'";

filters[0x2019] = "'";

filters[0x201C] = "\"";

filters[0x201D] = "\"";

filters[0x0007] = "\t";

filters[0x000D] = "\n";

filters[0x2002] = " ";

filters[0x2003] = " ";

filters[0x2012] = "-";

filters[0x2013] = "-";

filters[0x2014] = "-";

filters[0x000A] = "\n";

filters[0x000D] = "\n";

translations = [];

translations[0x82] = 0x201A;

translations[0x83] = 0x0192;

translations[0x84] = 0x201E;

translations[0x85] = 0x2026;

translations[0x86] = 0x2020;

translations[0x87] = 0x2021;

translations[0x88] = 0x02C6;

translations[0x89] = 0x2030;

translations[0x8A] = 0x0160;

translations[0x8B] = 0x2039;

translations[0x8C] = 0x0152;

translations[0x91] = 0x2018;

translations[0x92] = 0x2019;

translations[0x93] = 0x201C;

translations[0x94] = 0x201D;

translations[0x95] = 0x2022;

translations[0x96] = 0x2013;

translations[0x97] = 0x2014;

translations[0x98] = 0x02DC;

translations[0x99] = 0x2122;

translations[0x9A] = 0x0161;

translations[0x9B] = 0x203A;

translations[0x9C] = 0x0153;

translations[0x9F] = 0x0178;

Word = (function() {
  function Word(options) {
    if (typeof options === 'string') {
      options = {
        filename: options
      };
    } else {
      if (options == null) {
        options = {};
      }
    }
    this.filters = filters;
    this.fcTranslations = translations;
    this.filename = options.filename;
    console.log("Filename", this.filename);
  }

  Word.prototype.streamToBuffer = function(stream, cb) {
    var chunks;
    chunks = [];
    stream.on('data', function(chunk) {
      return chunks.push(chunk);
    });
    stream.on('error', function(error) {
      return cb(error);
    });
    return stream.on('end', function() {
      return cb(null, Buffer.concat(chunks));
    });
  };

  Word.prototype.extractBuffer = function(buffer, cb) {
    var flags, instance, magic, stream, table;
    instance = this;
    magic = buffer.readUInt16LE(0);
    if (magic !== 0xa5ec) {
      return cb("Invalid magic number");
    }
    flags = buffer.readUInt16LE(0xA);
    console.log("Flags", flags.toString(16));
    table = (flags & 0x0200) !== 0 ? "1Table" : "0Table";
    stream = this.doc.stream(table);
    return this.streamToBuffer(stream, (function(_this) {
      return function(error, data) {
        if (error != null) {
          return cb(error);
        }
        console.log("Read table buffer", data.length);
        _this.tableData = data;
        console.log("Table data", _this.tableData);
        _this.fcMin = _this.data.readUInt32LE(0x0018);
        _this.ccpText = _this.data.readUInt32LE(0x004c);
        _this.ccpFtn = _this.data.readUInt32LE(0x0050);
        _this.ccpHdd = _this.data.readUInt32LE(0x0054);
        _this.ccpAtn = _this.data.readUInt32LE(0x005c);
        _this.charPLC = _this.data.readUInt32LE(0x00fa);
        _this.charPlcSize = _this.data.readUInt32LE(0x00fe);
        _this.parPLC = _this.data.readUInt32LE(0x0102);
        _this.parPlcSize = _this.data.readUInt32LE(0x0106);
        _this.fcPlcfBteChpx = _this.data.readUInt32LE(0x0fa);
        _this.lcbPlcfBteChpx = _this.data.readUInt32LE(0x0fe);
        console.log("fcMin", _this.fcMin);
        console.log("ccpText", _this.ccpText);
        _this.getBookmarks();
        _this.pieces = _this.getPieces();
        return cb(null);
      };
    })(this));
  };

  Word.prototype.getBookmarks = function() {
    var bookmarks, cpEnd, cpStart, index, length, offset, segment;
    this.fcSttbfBkmk = this.data.readUInt32LE(0x0142);
    this.lcbSttbfBkmk = this.data.readUInt32LE(0x0146);
    this.fcPlcfBkf = this.data.readUInt32LE(0x014a);
    this.lcbPlcfBkf = this.data.readUInt32LE(0x014e);
    this.fcPlcfBkl = this.data.readUInt32LE(0x0152);
    this.lcbPlcfBkl = this.data.readUInt32LE(0x0156);
    if (this.lcbSttbfBkmk === 0) {
      return;
    }
    this.sttbfBkmk = this.tableData.slice(this.fcSttbfBkmk, this.fcSttbfBkmk + this.lcbSttbfBkmk);
    this.plcfBkf = this.tableData.slice(this.fcPlcfBkf, this.fcPlcfBkf + this.lcbPlcfBkf);
    this.plcfBkl = this.tableData.slice(this.fcPlcfBkl, this.fcPlcfBkl + this.lcbPlcfBkl);
    this.fcExtend = this.sttbfBkmk.readUInt16LE(0);
    this.cData = this.sttbfBkmk.readUInt16LE(2);
    this.cbExtra = this.sttbfBkmk.readUInt16LE(4);
    if (!($fcExtend === 0xffff)) {
      confess("Internal error: unexpected single-byte bookmark data");
    }
    offset = 6;
    index = 0;
    bookmarks = {};
    while (offset < this.lcbSttbfBkmk) {
      length = this.sttbfBkmk.readUInt16LE(offset);
      length = length * 2;
      segment = this.sttbfBkmk.slice(offset + 2, offset + 2 + length);
      cpStart = this.plcfBkf.readUInt32LE(index * 4);
      cpEnd = this.plcfBkl.readUInt32LE(index * 4);
      bookmarks[segment] = {
        start: cpStart,
        end: cpEnd
      };
      offset = offset + length + 2;
    }
    return this.bookmarks = bookmarks;
  };

  Word.prototype.filter = function(text, shouldFilter) {
    var matcher, replacer;
    if (!shouldFilter) {
      return text;
    }
    replacer = function() {
      var match, replaced, rest;
      match = arguments[0], rest = 2 <= arguments.length ? slice1.call(arguments, 1) : [];
      if (match.length === 1) {
        replaced = this.filters[match.charPointAt(0)];
        if (replaced === 0) {
          return "";
        } else {
          return replaced;
        }
      } else if (rest.length === 2) {
        return "";
      } else if (rest.length === 3) {
        return rest[0];
      }
    };
    matcher = /(?:[\x02\x05\x07\x08\x0a\x0d\u2018\u2019\u201c\u201d\u2002\u2003\u2012\u2013\u2014]|\x13(?:[^\x14]*\x14)?([^\x15]*)\x15)/g;
    return text.replace(replacer);
  };

  Word.prototype.getTextRange = function(start, end) {
    var endPiece, i, j, piece, pieces, ref, ref1, result, startPiece, xend, xstart;
    pieces = this.pieces;
    startPiece = this.getPieceIndex(start);
    endPiece = this.getPieceIndex(end);
    result = [];
    for (i = j = ref = startPiece, ref1 = endPiece; j <= ref1; i = j += 1) {
      piece = pieces[i];
      xstart = i === startPiece ? start - piece.position : 0;
      xend = i === endPiece ? end - piece.position : piece.endPosition;
      result.push(piece.text.substring(xstart, xend - xstart));
    }
    return result.join("");
  };

  Word.prototype.getPieceIndex = function(position) {
    var i, j, len, piece, ref;
    ref = this.pieces;
    for (i = j = 0, len = ref.length; j < len; i = ++j) {
      piece = ref[i];
      if (position <= piece.endPosition) {
        return i;
      }
    }
  };

  Word.prototype.getBody = function(shouldFilter) {
    var string;
    if (shouldFilter == null) {
      shouldFilter = true;
    }
    string = this.getTextRange(0, this.ccpText);
    return this.filter(string, shouldFilter);
  };

  Word.prototype.getText = function() {
    var index, j, len, length, piece, pieces, position, result, results, segment;
    result = [];
    index = 1;
    position = 0;
    pieces = this.pieces;
    results = [];
    for (j = 0, len = pieces.length; j < len; j++) {
      piece = pieces[j];
      piece.position = position;
      this.getPiece(piece);
      segment = piece.text;
      result.push(segment);
      length = segment.length;
      piece.length = length;
      piece.endPosition = position + length;
      results.push(position = position + length);
    }
    return results;
  };

  Word.prototype.getPieces = function() {
    var filePos, flag, j, lEnd, lStart, offset, piece, pieceTableSize, pieces, pos, ref, result, skip, start, totLength, unicode, x;
    pos = this.data.readUInt32LE(0x01a2);
    console.log("Pos", pos);
    result = [];
    while (true) {
      flag = this.tableData.readUInt8(pos);
      console.log("Flag skipping", flag);
      if (flag !== 1) {
        break;
      }
      pos = pos + 1;
      skip = this.tableData.readUInt16LE(pos);
      pos = pos + 2 + skip;
      console.log("Skipping", skip);
    }
    flag = this.tableData.readUInt8(pos);
    console.log("Flag after skipping", flag);
    pos = pos + 1;
    if (flag !== 2) {
      throw new Error("Internal error: ccorrupted Word file");
    }
    pieceTableSize = this.tableData.readUInt32LE(pos);
    pos = pos + 4;
    pieces = (pieceTableSize - 4) / 12;
    start = 0;
    for (x = j = 0, ref = pieces - 1; j <= ref; x = j += 1) {
      offset = pos + ((pieces + 1) * 4) + (x * 8) + 2;
      filePos = this.tableData.readUInt32LE(offset);
      unicode = false;
      if ((filePos & 0x40000000) === 0) {
        unicode = true;
      } else {
        filePos = filePos & ~0x40000000;
        filePos = Math.floor(filePos / 2);
      }
      lStart = this.tableData.readUInt32LE(pos + (x * 4));
      lEnd = this.tableData.readUInt32LE(pos + ((x + 1) * 4));
      totLength = lEnd - lStart;
      console.log("lStart", lStart, "lEnd", lEnd);
      piece = {
        start: start,
        totLength: totLength,
        filePos: filePos,
        unicode: unicode
      };
      console.log("Piece", piece);
      result.push(piece);
      start = start + (unicode ? Math.floor(totLength / 2) : totLength);
    }
    return result;
  };

  Word.prototype.getPiece = function(piece) {
    var pend, pfilePos, pstart, ptotLength, punicode, textEnd, textStart;
    pstart = piece.start;
    ptotLength = piece.totLength;
    pfilePos = piece.filePos;
    punicode = piece.unicode;
    pend = pstart + ptotLength;
    textStart = pfilePos;
    textEnd = textStart + (pend - pstart);
    if (punicode) {
      return piece.text = this.addUnicodeText(textStart, textEnd);
    } else {
      return piece.text = this.addText(textStart, textEnd);
    }
  };

  Word.prototype.addText = function(textStart, textEnd) {
    var slice;
    console.log("Adding text", textStart, textEnd);
    slice = this.data.slice(textStart, textEnd);
    return slice.toString('binary');
  };

  Word.prototype.addUnicodeText = function(textStart, textEnd) {
    var slice, string;
    slice = this.data.slice(textStart, 2 * textEnd - textStart);
    string = slice.toString('ucs2');
    return string;
  };

  Word.prototype.extract = function(cb) {
    console.log("File", this.filename);
    this.doc = new oleDoc(this.filename);
    this.doc.on('err', (function(_this) {
      return function(err) {
        console.log("Error", err);
        return cb(err);
      };
    })(this));
    this.doc.on('ready', (function(_this) {
      return function() {
        var stream;
        stream = _this.doc.stream('WordDocument');
        return _this.streamToBuffer(stream, function(error, data) {
          if (error != null) {
            return cb(error);
          }
          _this.data = data;
          return _this.extractBuffer(data, function(error) {
            return cb(error);
          });
        });
      };
    })(this));
    return this.doc.read();
  };

  return Word;

})();

module.exports = Word;
