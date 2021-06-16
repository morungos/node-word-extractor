const path = require('path');
const WordExtractor = require('../lib/word');

describe('Word file bad-xml.docx', () => {

  const extractor = new WordExtractor();

  it.skip('should extract the expected body', () => {
    return extractor.extract(path.resolve(__dirname, "data/bad-xml.docx"))
      .then((document) => {
        const value = JSON.stringify(document.getBody());
        expect(value).toBeInstanceOf(String);
      });
  });
});
