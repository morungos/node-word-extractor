const path = require('path');
const WordExtractor = require('../lib/word');

describe('Word file test03.doc', () => {

  const extractor = new WordExtractor();

  it('should match the expected body', (done) => {
    const extract = extractor.extract(path.resolve(__dirname, "data/test03.doc"));
    return extract
      .then((result) => {
        const body = result.getBody();
        expect(body).toMatch(/Can You Release Commercial Works\?/);
        expect(body).toMatch(/Apache v2\.0/);
        expect(body).toMatch(/you want your protections to be\./);
        done();
      });
  });
});
