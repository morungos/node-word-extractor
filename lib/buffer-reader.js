class BufferReader {

  constructor(buffer) {
    this._buffer = buffer;
  }

  open() {
    return Promise.resolve();
  }

  close() {
    return Promise.resolve();
  }

  read(buffer, offset, length, position) {
    this._buffer.copy(buffer, offset, position, position + length);
    return Promise.resolve(buffer);
  }
}

module.exports = BufferReader;
