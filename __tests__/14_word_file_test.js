const path = require('path');
const WordExtractor = require('../lib/word');

describe('Word file test04.doc', () => {

  const extractor = new WordExtractor();

  it('should match the expected body', (done) => {
    const extract = extractor.extract(path.resolve(__dirname, "data/test04.doc"));
    return extract
      .then((result) => {
        const body = result.getBody();
        expect(body).toMatch(/Moli\u00e8re/);
        done();
    });
  });

  it('should have headers and footers', (done) => {
    const extract = extractor.extract(path.resolve(__dirname, "data/test04.doc"));
    return extract
      .then((result) => {
        const body = result.getHeaders();
        expect(body).toMatch(/The footer/);
        expect(body).toMatch(/Moli\u00e8re/);
        done();
    });
  });
});
