const path = require('path');
const WordExtractor = require('../lib/word');

describe('Word file test15.doc', () => {

  const extractor = new WordExtractor();

  it('should match the expected body', () => {
    const extract = extractor.extract(path.resolve(__dirname, "data/test15.doc"));
    return extract
      .then((result) => {
        const body = result.getBody();
        expect(body).toMatch(/Endnotes and footnotes test\n\nParagraph 1\n\nParagraph 2/);

        const footnotes = result.getFootnotes();
        expect(footnotes).toMatch(/This is a footnote/);

        const endnotes = result.getEndnotes();
        expect(endnotes).toMatch(/This is an endnote/);
      });
  });
});
