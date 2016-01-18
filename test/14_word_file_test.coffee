chai = require('chai')
expect = chai.expect

fs = require('fs')
path = require('path')
Buffer = require('buffer').Buffer

WordExtractor = require('../src/word')
Document = require('../src/document')

describe 'Word file test04.doc', () ->

  extractor = new WordExtractor()

  it 'should match the expected body', (done) ->
    extract = extractor.extract path.resolve(__dirname, "data/test04.doc")
    expect(extract).to.be.fulfilled
    extract
      .then (result) ->
        body = result.getBody()
        expect(body).to.match new RegExp('Moli\\u00e8re')
        done()

  it 'should have headers and footers', (done) ->
    extract = extractor.extract path.resolve(__dirname, "data/test04.doc")
    expect(extract).to.be.fulfilled
    extract
      .then (result) ->
        body = result.getHeaders()
        expect(body).to.match new RegExp('The footer')
        expect(body).to.match new RegExp('Moli\\u00e8re')
        done()
