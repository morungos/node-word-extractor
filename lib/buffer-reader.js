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

  buffer() {
    return this._buffer;
  }

  static isBufferReader(instance) {
    return instance instanceof BufferReader;
  }
}

module.exports = BufferReader;
