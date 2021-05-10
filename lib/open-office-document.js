class OpenOfficeDocument {

  constructor() {
    this._segments = [];
    this._index = 0;
  }

  add(string) {
    this._segments.push(string);
    this._index += string.length;
  }

  getBody() {
    return this._segments.join("");
  }

}


module.exports = OpenOfficeDocument;
