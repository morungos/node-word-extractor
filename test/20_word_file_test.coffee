chai = require('chai')
expect = chai.expect

fs = require('fs')
path = require('path')
Buffer = require('buffer').Buffer

WordExtractor = require('../lib/word')
Document = require('../lib/document')

describe 'Word file test10.doc', () ->

  extractor = new WordExtractor()

  it 'should match the expected body', (done) ->
    extract = extractor.extract path.resolve(__dirname, "data/test10.doc")
    expect(extract).to.be.fulfilled
    extract
      .then (result) ->
        body = result.getBody()
        expect(body).to.match new RegExp('Second paragraph')
        done()

  it 'should retrieve the annotations', (done) ->
    extract = extractor.extract path.resolve(__dirname, "data/test10.doc")
    expect(extract).to.be.fulfilled
    extract
      .then (result) ->
        annotations = result.getAnnotations()
        expect(annotations).to.match new RegExp('Second paragraph')
        done()
