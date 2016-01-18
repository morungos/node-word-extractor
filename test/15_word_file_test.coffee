chai = require('chai')
expect = chai.expect

fs = require('fs')
path = require('path')
Buffer = require('buffer').Buffer

WordExtractor = require('../src/word')
Document = require('../src/document')

describe 'Word file test05.doc', () ->

  extractor = new WordExtractor()

  it 'should match the expected body', (done) ->
    extract = extractor.extract path.resolve(__dirname, "data/test05.doc")
    expect(extract).to.be.fulfilled
    extract
      .then (result) ->
        body = result.getBody()
        expect(body).to.match new RegExp('This is a simple file created with Word 97-SR2')
        done()
