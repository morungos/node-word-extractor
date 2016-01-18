chai = require('chai')
expect = chai.expect

fs = require('fs')
path = require('path')
Buffer = require('buffer').Buffer

WordExtractor = require('../src/word')
Document = require('../src/document')

testWordFile = (file) ->

  describe 'Word file ' + file, () ->

    extractor = new WordExtractor()

    it 'should extract a document successfully', (done) ->
      extract = extractor.extract path.resolve(__dirname, "data/" + file)
      expect(extract).to.be.fulfilled
      extract
        .then (result) ->
          expect(result).to.be.an.instanceof(Document)
          expect(result.pieces).to.be.instanceof(Array)
          expect(result.boundaries).to.contain.keys(['fcMin', 'ccpText', 'ccpFtn', 'ccpHdd', 'ccpAtn'])
          done()

files = fs.readdirSync(path.resolve(__dirname, "data"))
files.filter (file) ->
  testWordFile(file) if /\.doc$/i.test(file)
