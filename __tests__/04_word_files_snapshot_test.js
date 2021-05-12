/**
 * @overview
 * Snapshot tests for all Word (.doc) files. The useful thing about
 * this is it detects changes, but also the snapshots include the binary
 * values and characters, so we see exactly what is returned, which is
 * extremely useful for debugging.
 */

const fs = require('fs');
const path = require('path');
const WordExtractor = require('../lib/word');

const files = fs.readdirSync(path.resolve(__dirname, "data"));
const pairs = files.filter((f) => f.match(/test(\d+)\.doc$/))
  .filter((f) => ! /^~/.test(f));

describe.each(pairs.map((x) => [x]))(
  `Word file %s`, (file) => {

    const extractor = new WordExtractor();

    it('should match its snapshot', function() {
      return extractor.extract(path.resolve(__dirname, `data/${file}`))
        .then((document) => {
          const value = JSON.stringify({
            body: document.getBody(),
            footnotes: document.getFootnotes(),
            endnotes: document.getEndnotes(),
            headers: document.getHeaders(),
            annotations: document.getAnnotations(),
          });
          expect(value).toMatchSnapshot();
        });
    });
  }
);
