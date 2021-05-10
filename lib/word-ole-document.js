const { filter } =        require('./filters');

const getPieceIndex = (pieces, position) => {
  for (let i = 0; i < pieces.length; i++) {
    const piece = pieces[i];
    if (position <= piece.endPosition) { 
      return i; 
    }
  }
};

class WordOleDocument {

  constructor() {
    this.pieces = [];
    this.bookmarks = {};
    this.boundaries = {};
  }

  getTextRange(start, end) {
    const { pieces } = this;
    const startPiece = getPieceIndex(pieces, start);
    const endPiece = getPieceIndex(pieces, end);
    const result = [];
    for (let i = startPiece, end1 = endPiece; i <= end1; i++) {
      const piece = pieces[i];
      const xstart = i === startPiece ? start - piece.position : 0;
      const xend = i === endPiece ? end - piece.position : piece.endPosition;
      result.push(piece.text.substring(xstart, xend));
    }

    return result.join("");
  }

  getBody() {
    const start = 0;
    const string = this.getTextRange(start, start + this.boundaries.ccpText);
    return filter(string);
  }


  getFootnotes() {
    const start = this.boundaries.ccpText;
    const string = this.getTextRange(start, start + this.boundaries.ccpFtn);
    return filter(string);
  }


  getHeaders() {
    const start = this.boundaries.ccpText + this.boundaries.ccpFtn;
    const string = this.getTextRange(start, start + this.boundaries.ccpHdd);
    return filter(string);
  }


  getAnnotations() {
    const start = this.boundaries.ccpText + this.boundaries.ccpFtn + this.boundaries.ccpHdd;
    const string = this.getTextRange(start, start + this.boundaries.ccpAtn);
    return filter(string);
  }


  getEndnotes() {
    const start = this.boundaries.ccpText + this.boundaries.ccpFtn + this.boundaries.ccpHdd + this.boundaries.ccpAtn;
    const string = this.getTextRange(start, start + this.boundaries.ccpAtn + this.boundaries.ccpEdn);
    return filter(string);
  }
}


module.exports = WordOleDocument;
