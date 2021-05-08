const { Buffer } =       require('buffer');

const WordOleExtractor = require('./word-ole-extractor');

const BufferReader =     require('./buffer-reader');
const FileReader =       require('./file-reader');

class WordExtractor {

  constructor() {}

  /**
   * Extracts the main contents of the file. If a Buffer is passed, that
   * is used instead.
   * @param {*} source either a string filename, or a Buffer containing the file content
   * @returns a {@link Document} providing accessors onto the text
   */
  extract(source) {
    let reader = null;
    if (Buffer.isBuffer(source)) {
      reader = new BufferReader(source);
    } else if (typeof source === 'string') {
      reader = new FileReader(source);
    }
    const buffer = Buffer.alloc(512);
    return reader.open()
      .then(() => reader.read(buffer, 0, 512, 0))
      .then((buffer) => {
        let extractor = null;

        if (buffer.readUInt16BE(0) === 0xd0cf) {
          extractor = WordOleExtractor;
        } else if (buffer.readUInt16BE(0) === 0x504b) {
          const next = buffer.readUInt16BE(2);
          if ((next === 0x0304) || (next === 0x0506) || (next === 0x0708)) {
            throw new Error("Extracting from .docx files not yet implemented");
          }
        }

        if (! extractor) {
          throw new Error("Unable to read this type of file");
        }

        return (new extractor()).extract(reader);
      })
      .finally(() => reader.close());
  }

}


module.exports = WordExtractor;
