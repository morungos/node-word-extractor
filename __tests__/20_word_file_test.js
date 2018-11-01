const path = require('path');
const WordExtractor = require('../lib/word');

describe('Word file test10.doc', () => {

  const extractor = new WordExtractor();

  it('should match the expected body', (done) => {
    const extract = extractor.extract(path.resolve(__dirname, "data/test10.doc"));
    return extract
      .then((result) => {
        const body = result.getBody();
        expect(body).toMatch(/Second paragraph/);
        done();
      })
      .catch(err => done(err));
  });

  it('should retrieve the annotations', (done) => {
    const extract = extractor.extract(path.resolve(__dirname, "data/test10.doc"));
    return extract
      .then((result) => {
        const annotations = result.getAnnotations();
        expect(annotations).toMatch(/Second paragraph comment/);
        expect(annotations).toMatch(/Third paragraph comment/);
        expect(annotations).toMatch(/this is all I have to say on the matter/);
        done();
      })
      .catch(err => done(err));
  });
});
