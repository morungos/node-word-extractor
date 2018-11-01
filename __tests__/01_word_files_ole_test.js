const fs = require('fs');
const path = require('path');
const { Buffer } = require('buffer');

const oleDoc = require('../lib/ole-doc').OleCompoundDoc;

const files = fs.readdirSync(path.resolve(__dirname, "data"));
describe.each(files.filter((f) => f.match(/\.doc$/)).map((x) => [x]))(
  `Word file %s`, (file) => {
    it('can be opened correctly', (done) => {
      const filename = path.resolve(__dirname, `data/${file}`);
      const doc = new oleDoc(filename);
      doc.on('err', () => done(err));
      doc.on('ready', () => done());
      doc.read();
    });

    it('generates a valid Word stream', (done) => {
      const filename = path.resolve(__dirname, `data/${file}`);
      const doc = new oleDoc(filename);

      doc.on('err', err => done(err));
      doc.on('ready', () => {
        const chunks = [];
        const stream = doc.stream('WordDocument');
        stream.on('data', chunk => chunks.push(chunk));
        stream.on('error', error => done(error));
        stream.on('end', () => {
          const buffer = Buffer.concat(chunks);
          const magicNumber = buffer.readUInt16LE(0);
          expect(magicNumber.toString(16)).toBe("a5ec");
          done();
        });
      });
      doc.read();
    });
  }
);
