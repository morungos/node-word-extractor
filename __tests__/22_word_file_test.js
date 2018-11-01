const path = require('path');
const WordExtractor = require('../lib/word');

describe('Word file test12.doc', () => {

  const extractor = new WordExtractor();

  it('should match the expected body', () =>
    extractor.extract(path.resolve(__dirname, "data/test12.doc"))
      .then((result) => {
        const body = result.getBody();
        expect(body).toMatch(/Row 2, cell 3/);
    })
  );
});
