const filters =        require('./filters');

const getPieceIndex = (pieces, position) => {
  for (let i = 0; i < pieces.length; i++) {
    const piece = pieces[i];
    if (position <= piece.endPosition) { 
      return i; 
    }
  }
};

const filter = (text, shouldFilter) => {
  
  if (!shouldFilter) { 
    return text; 
  }

  const replacer = function(match, ...rest) {
    if (match.length === 1) {
      const replaced = filters[match.charCodeAt(0)];
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

  const matcher = /(?:[\x02\x05\x07\x08\x0a\x0d\u2018\u2019\u201c\u201d\u2002\u2003\u2012\u2013\u2014]|\x13(?:[^\x14]*\x14)?([^\x15]*)\x15)/g;
  return text.replace(matcher, replacer);
};

class Document {

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

  getBody(shouldFilter) {
    if (shouldFilter == null) { 
      shouldFilter = true; 
    }
    const start = 0;
    const string = this.getTextRange(start, start + this.boundaries.ccpText);
    return filter(string, shouldFilter);
  }


  getFootnotes(shouldFilter) {
    if (shouldFilter == null) { 
      shouldFilter = true; 
    }
    const start = this.boundaries.ccpText;
    const string = this.getTextRange(start, start + this.boundaries.ccpFtn);
    return filter(string, shouldFilter);
  }


  getHeaders(shouldFilter) {
    if (shouldFilter == null) { 
      shouldFilter = true; 
    }
    const start = this.boundaries.ccpText + this.boundaries.ccpFtn;
    const string = this.getTextRange(start, start + this.boundaries.ccpHdd);
    return filter(string, shouldFilter);
  }


  getAnnotations(shouldFilter) {
    if (shouldFilter == null) { 
      shouldFilter = true; 
    }
    const start = this.boundaries.ccpText + this.boundaries.ccpFtn + this.boundaries.ccpHdd;
    const string = this.getTextRange(start, start + this.boundaries.ccpAtn);
    return filter(string, shouldFilter);
  }


  getEndnotes(shouldFilter) {
    if (shouldFilter == null) { 
      shouldFilter = true; 
    }
    const start = this.boundaries.ccpText + this.boundaries.ccpFtn + this.boundaries.ccpHdd + this.boundaries.ccpAtn;
    const string = this.getTextRange(start, start + this.boundaries.ccpAtn + this.boundaries.ccpEdn);
    return filter(string, shouldFilter);
  }
}


module.exports = Document;
