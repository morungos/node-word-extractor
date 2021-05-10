const fs = require('fs');
const path = require('path');
const { Buffer } = require('buffer');

const OleCompoundDoc = require('../lib/ole-compound-doc');
const FileReader = require('../lib/file-reader');
const BufferReader = require('../lib/buffer-reader');

const files = fs.readdirSync(path.resolve(__dirname, "data"))
  .filter((f) => ! /^~/.test(f))
  .filter((f) => f.match(/test(\d+)\.doc$/));

describe.each(files.map((x) => [x]))(
  `Word file %s`, (file) => {
    it('can be opened correctly', () => {
      const filename = path.resolve(__dirname, `data/${file}`);
      const reader = new FileReader(filename);
      const doc = new OleCompoundDoc(reader);
      return reader.open()
        .then(() => doc.read())
        .finally(() => reader.close());
    });

    it('generates a valid Word stream', () => {
      const filename = path.resolve(__dirname, `data/${file}`);
      const reader = new FileReader(filename);
      const doc = new OleCompoundDoc(reader);

      return reader.open()
        .then(() => doc.read())
        .then(() => {
          return new Promise((resolve, reject) => {
            const chunks = [];
            const stream = doc.stream('WordDocument');
            stream.on('data', (chunk) => chunks.push(chunk));
            stream.on('error', (error) => reject(error));
            stream.on('end', () => resolve(Buffer.concat(chunks)));
          });
        })
        .then((buffer) => {
          const magicNumber = buffer.readUInt16LE(0);
          expect(magicNumber.toString(16)).toBe("a5ec");
        })
        .finally(() => reader.close());
    });

    it('generates a valid Word stream from a buffer', () => {
      const filename = path.resolve(__dirname, `data/${file}`);
      const buffer = fs.readFileSync(filename);
      const reader = new BufferReader(buffer);
      const doc = new OleCompoundDoc(reader);

      return reader.open()
        .then(() => doc.read())
        .then(() => {
          return new Promise((resolve, reject) => {
            const chunks = [];
            const stream = doc.stream('WordDocument');
            stream.on('data', (chunk) => chunks.push(chunk));
            stream.on('error', (error) => reject(error));
            stream.on('end', () => resolve(Buffer.concat(chunks)));
          });
        })
        .then((buffer) => {
          const magicNumber = buffer.readUInt16LE(0);
          expect(magicNumber.toString(16)).toBe("a5ec");
        })
        .finally(() => reader.close());
    });

  }
);
