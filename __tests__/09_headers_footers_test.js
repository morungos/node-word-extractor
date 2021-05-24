const path = require('path');
const WordExtractor = require('../lib/word');

describe('Word file word15.docx', () => {

  const extractor = new WordExtractor();

  it('should properly separate headers and footers', () => {
    return extractor.extract(path.resolve(__dirname, "data/test15.docx"))
      .then((document) => {
        expect(document.getFooters()).toMatch(/footer/);
        expect(document.getFooters()).not.toMatch(/header/);
        expect(document.getHeaders({includeFooters: false})).toMatch(/header/);
        expect(document.getHeaders({includeFooters: false})).not.toMatch(/footer/);
      });
  });
});


describe('Word file word15.doc', () => {

  const extractor = new WordExtractor();

  it('should properly separate headers and footers', () => {
    return extractor.extract(path.resolve(__dirname, "data/test15.doc"))
      .then((document) => {
        expect(document.getFooters()).toMatch(/footer/);
        expect(document.getFooters()).not.toMatch(/header/);
        expect(document.getHeaders({includeFooters: false})).toMatch(/header/);
        expect(document.getHeaders({includeFooters: false})).not.toMatch(/footer/);
      });
  });
});
