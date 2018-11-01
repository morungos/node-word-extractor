const fs = require('fs');
const path = require('path');
const { Buffer } = require('buffer');

const WordExtractor = require('../lib/index');

describe('Checking block from files', () => {

  const extractor = new WordExtractor();

  it('should extract a .doc document successfully', () =>
    extractor.extract(path.resolve(__dirname, "data/test01.doc"))
  );

  it('should extract a .docx document successfully', () =>
    extractor.extract(path.resolve(__dirname, "data/test13.docx"))
  );

});
