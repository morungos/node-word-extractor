const path = require('path');
const WordExtractor = require('../lib/index');

describe('Checking block from files', () => {

  const extractor = new WordExtractor();

  it('should extract a .doc document successfully', () => {
    return extractor.extract(path.resolve(__dirname, "data/test01.doc"));
  });

  it('should extract a .docx document successfully', () => {
    return extractor.extract(path.resolve(__dirname, "data/test13.docx"));
  });

  it('should handle missing file error correctly', () => {
    const result = extractor.extract(path.resolve(__dirname, "data/missing00.docx"));
    return expect(result).rejects.toEqual(expect.objectContaining({
      message: expect.stringMatching(/no such file or directory/)
    }));
  });

});
 