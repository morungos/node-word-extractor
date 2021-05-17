const path = require('path');
const WordExtractor = require('../lib/word');

describe('Word file badfile-01-bad-header.doc', () => {

  const extractor = new WordExtractor();

  it('should match the expected body', () => {
    return expect(extractor.extract(path.resolve(__dirname, "data/badfile-01-bad-header.doc")))
      .rejects
      .toThrowError("Invalid Short Sector Allocation Table");
  });
});
