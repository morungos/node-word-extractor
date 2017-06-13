## Early tests, which handle some of the type detection.

chai = require('chai')
expect = chai.expect

fs = require('fs')
path = require('path')
Buffer = require('buffer').Buffer

WordExtractor = require('../src/index')

describe 'Checking block from files', () ->

  extractor = new WordExtractor()

  it 'should extract a .doc document successfully', (done) ->
    extractor.extract path.resolve(__dirname, "data/test01.doc")
      .then () ->
        done()

  it 'should extract a .docx document successfully', (done) ->
    extractor.extract path.resolve(__dirname, "data/test13.docx")
      .then () ->
        done()
