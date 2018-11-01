const path = require('path');
const WordExtractor = require('../lib/word');

describe('Word file test06.doc', () => {

  const extractor = new WordExtractor();

  it('should match the expected body', (done) => {
    const extract = extractor.extract(path.resolve(__dirname, "data/test06.doc"));
    return extract
      .then((result) => {
        const body = result.getBody();
        expect(body).toMatch(/Insert interface name here/);
        expect(body).toMatch(/M\u00e9thodes appel\u00e9es/);
        done();
    });
  });
});
