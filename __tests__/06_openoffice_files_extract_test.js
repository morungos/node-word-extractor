/**
 * @overview
 * Snapshot tests for all Office Open XML (.docx) files. The useful thing about
 * this is it detects changes, but also the snapshots include the binary values
 * and characters, so we see exactly what is returned, which is extremely
 * useful for debugging.
 */

const fs = require('fs');
const path = require('path');
const WordExtractor = require('../lib/word');

require('jest-specific-snapshot');

const files = fs.readdirSync(path.resolve(__dirname, "data"));
const pairs = files.filter((f) => f.match(/test(\d+)\.doc$/))
  .filter((f) => files.includes(f + "x"))
  .filter((f) => ! /^~/.test(f));

const cleanHeaderText = (text) => {
  return text.replace(/^\s+/, '')
    .replace(/\s+$/, '')
    .replace(/\n{2,}/g, '\n\n');
};

describe.each(pairs.map((x) => [x]))(
  `Word file %s`, (file) => {

    const extractor = new WordExtractor();

    it('should match its snapshot', () => {
      return extractor.extract(path.resolve(__dirname, `data/${file}x`))
        .then((document) => {
          const value = {
            body: JSON.stringify(document.getBody()),
            footnotes: JSON.stringify(document.getFootnotes()),
            endnotes: JSON.stringify(document.getEndnotes()),
            headers: JSON.stringify(document.getHeaders()),
            annotations: JSON.stringify(document.getAnnotations()),
            textboxes: JSON.stringify(document.getTextboxes({includeHeadersAndFooters: false}).trim()),
            headerTextboxes: JSON.stringify(cleanHeaderText(document.getTextboxes({includeBody: false})).trim())
          };
          expect(value).toMatchSpecificSnapshot(`./__snapshots__/${file}.snapx`, {
            headers: expect.any(String)
          });
        });
    });
  }
);
