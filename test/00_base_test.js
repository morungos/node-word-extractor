/*
 * decaffeinate suggestions:
 * DS102: Remove unnecessary code created because of implicit returns
 * Full docs: https://github.com/decaffeinate/decaffeinate/blob/master/docs/suggestions.md
 */
//# Early tests, which handle some of the type detection.

const chai = require('chai');
const { expect } = chai;

const fs = require('fs');
const path = require('path');
const { Buffer } = require('buffer');

const WordExtractor = require('../lib/index');

describe('Checking block from files', function() {

  const extractor = new WordExtractor();

  it('should extract a .doc document successfully', done =>
    extractor.extract(path.resolve(__dirname, "data/test01.doc"))
      .then(() => done())
  );

  return it('should extract a .docx document successfully', done =>
    extractor.extract(path.resolve(__dirname, "data/test13.docx"))
      .then(() => done())
  );
});
