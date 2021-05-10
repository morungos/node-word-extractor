const fs = require('fs');
const path = require('path');
const WordExtractor = require('../lib/word');

describe('Checking block from files', () => {

  const extractor = new WordExtractor();

  it('should extract a .doc document successfully', () => {
    return extractor.extract(path.resolve(__dirname, "data/test01.doc"));
  });

  it('should extract a .docx document successfully', () => {
    return extractor.extract(path.resolve(__dirname, "data/test01.docx"));
  });

  it('should handle missing file error correctly', () => {
    const result = extractor.extract(path.resolve(__dirname, "data/missing00.docx"));
    return expect(result).rejects.toEqual(expect.objectContaining({
      message: expect.stringMatching(/no such file or directory/)
    }));
  });

  it('should properly close the file', () => {
    const open = jest.spyOn(fs, 'open');
    const close = jest.spyOn(fs, 'close');
    return extractor.extract(path.resolve(__dirname, "data/test01.doc"))
      .then(() => {
        expect(open).toHaveBeenCalledTimes(1);
        expect(close).toHaveBeenCalledTimes(1);
      })
      .finally(() => {
        open.mockRestore();
        close.mockRestore();
      });
  });

});
 