const path = require('path');
const WordExtractor = require('../lib/word');

describe('Word file bad-xml.docx', () => {

  const extractor = new WordExtractor();

  it('should extract the expected body', () => {
    return extractor.extract(path.resolve(__dirname, "data/bad-xml.docx"))
      .then((document) => {
        expect(document.getBody()).toEqual(expect.stringMatching(/A second test of reviewing, but with Unicode characters in/));
      });
  });
});
