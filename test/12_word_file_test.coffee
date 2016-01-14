chai = require('chai')
expect = chai.expect

fs = require('fs')
path = require('path')
Buffer = require('buffer').Buffer

WordExtractor = require('../lib/word')
Document = require('../lib/document')

describe 'Word file test2.doc', () ->

  extractor = new WordExtractor()

  it 'should match the expected body', (done) ->
    extract = extractor.extract path.resolve(__dirname, "data/test2.doc")
    expect(extract).to.be.fulfilled
    extract
      .then (result) ->
        body = result.getBody()
        expect(body).to.match new RegExp('My name is Ryan')
        expect(body).to.match new RegExp('create several FKPs for testing purposes')
        expect(body).to.match new RegExp('dsadasdasdasd')
        done()
