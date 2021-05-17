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
const pairs = files.filter((f) => f.match(/bigfile-(\d+)\.doc$/))
  .filter((f) => files.includes(f + "x"))
  .filter((f) => ! /^~/.test(f));

describe.each(pairs.map((x) => [x]))(
  `Word file %s`, (file) => {

    const extractor = new WordExtractor();

    it('should match across formats', () => {
      return Promise.all([
        extractor.extract(path.resolve(__dirname, `data/${file}`)),
        extractor.extract(path.resolve(__dirname, `data/${file}x`))
      ])
        .then((documents) => {
          const [oleDocument, ooDocument] = documents;
          const oleBody = oleDocument.getBody().replace(/\n{2,}/g, "\n");
          const ooBody = ooDocument.getBody().replace(/\n{2,}/g, "\n");
          expect(oleBody).toEqual(ooBody);
        });
    });
  }
);
