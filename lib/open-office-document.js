class OpenOfficeDocument {

  constructor() {
    this._segments = [];
    this._index = 0;
  }

  add(string) {
    this._segments.push(string);
    this._index += string.length;
  }

}


module.exports = OpenOfficeDocument;
