const path = require('path');
const WordExtractor = require('../lib/word');

describe('Word file test09.doc', () => {

  const extractor = new WordExtractor();

  it('should match the expected body', (done) => {
    const extract = extractor.extract(path.resolve(__dirname, "data/test09.doc"));
    return extract
      .then((result) => {
        const body = result.getBody();
        expect(body).toMatch(/This line gets read fine/);
        expect(body).toMatch(/Ooops, where are the \( opening \( brackets\?/);
        done();
      });
  });
});
