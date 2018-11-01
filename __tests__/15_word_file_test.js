const path = require('path');
const WordExtractor = require('../lib/word');

describe('Word file test05.doc', () => {

  const extractor = new WordExtractor();

  it('should match the expected body', (done) => {
    const extract = extractor.extract(path.resolve(__dirname, "data/test05.doc"));
    return extract
      .then(function(result) {
        const body = result.getBody();
        expect(body).toMatch(/This is a simple file created with Word 97-SR2/);
        done();
      });
  });
});
