const fs = require('fs');
const path = require('path');
const WordExtractor = require('../lib/word');
const Document = require('../lib/document');

const files = fs.readdirSync(path.resolve(__dirname, "data"));
describe.each(files.filter((f) => f.match(/\.doc$/)).map((x) => [x]))(
  `Word file %s`, (file) => {

    const extractor = new WordExtractor();

    it('should extract a document successfully', function(done) {
      const extract = extractor.extract(path.resolve(__dirname, `data/${file}`));
      return extract
        .then(function(result) {
          expect(result).toBeInstanceOf(Document);
          expect(result.pieces).toBeInstanceOf(Array);
          expect(result.boundaries).toEqual(expect.objectContaining({
            'fcMin': expect.any(Number),
            'ccpText': expect.any(Number), 
            'ccpFtn': expect.any(Number), 
            'ccpHdd': expect.any(Number), 
            'ccpAtn': expect.any(Number)
          }));
          done();
        });
    });
  }
);
