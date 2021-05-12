const fs = require('fs');
const path = require('path');
const WordExtractor = require('../lib/word');
const Document = require('../lib/document');

const files = fs.readdirSync(path.resolve(__dirname, "data"))
  .filter((f) => ! /^~/.test(f))
  .filter((f) => f.match(/test(\d+)\.doc$/));

describe.each(files.map((x) => [x]))(
  `Word file %s`, (file) => {

    const extractor = new WordExtractor();

    it('should extract a document successfully', function() {
      const extract = extractor.extract(path.resolve(__dirname, `data/${file}`));
      return extract
        .then(function(result) {
          expect(result).toBeInstanceOf(Document);
        });
    });
  }
);
