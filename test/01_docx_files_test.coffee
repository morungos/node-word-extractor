chai = require('chai')
expect = chai.expect

fs = require('fs')
path = require('path')
Buffer = require('buffer').Buffer

DOCXExtractor = require('../src/docx');

testWordFile = (file) ->
  describe 'Word file ' + file, () ->

    it 'can be opened correctly', () ->
      filename = path.resolve(__dirname, "data/" + file)
      doc = new DOCXExtractor()
      doc.extract(filename)


files = fs.readdirSync(path.resolve(__dirname, "data"))
files.filter (file) ->
  testWordFile(file) if /\.docx$/i.test(file)
