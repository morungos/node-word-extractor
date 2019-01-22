const path = require('path');
const WordExtractor = require('../lib/word');

describe('Word file test11.doc', () => {

  const extractor = new WordExtractor();

  it('should match the expected body', (done) => {
    const extract = extractor.extract(path.resolve(__dirname, "data/test11.doc"));
    return extract
      .then((result) => {
        const body = result.getBody();
        expect(body).toMatch(/This is a test for parsing the Word file in node/);
        done();
      })
      .catch(err => done(err));
  });
});
