chai = require('chai')
expect = chai.expect

fs = require('fs')
path = require('path')
Buffer = require('buffer').Buffer

WordExtractor = require('../lib/word')
Document = require('../lib/document')

describe 'Word file test1.doc', () ->

  extractor = new WordExtractor()

  it 'should match the expected body', (done) ->
    extract = extractor.extract path.resolve(__dirname, "data/test1.doc")
    expect(extract).to.be.fulfilled
    extract
      .then (result) ->
        body = result.getBody()
        expect(body).to.match new RegExp('Welcome to BlogCFC')
        expect(body).to.match new RegExp('http://lyla\\.maestropublishing\\.com/')
        expect(body).to.match new RegExp('You must use the IDs\\.')
        done()
