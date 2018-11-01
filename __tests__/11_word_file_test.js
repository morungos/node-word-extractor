const fs = require('fs');
const path = require('path');
const { Buffer } = require('buffer');

const WordExtractor = require('../lib/word');
const Document = require('../lib/document');

describe('Word file test01.doc', () => {

  const extractor = new WordExtractor();

  it('should match the expected body', (done) => {
    const extract = extractor.extract(path.resolve(__dirname, "data/test01.doc"));
    return extract
      .then(function(result) {
        const body = result.getBody();
        expect(body).toMatch(/Welcome to BlogCFC/);
        expect(body).toMatch(/http:\/\/lyla\.maestropublishing\.com\//);
        expect(body).toMatch(/You must use the IDs\./);
        done();
    });
  });
});
