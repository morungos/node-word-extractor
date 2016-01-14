chai = require('chai')
expect = chai.expect

fs = require('fs')
path = require('path')
Buffer = require('buffer').Buffer

WordExtractor = require('../lib/word')
Document = require('../lib/document')

describe 'Word file test3.doc', () ->

  extractor = new WordExtractor()

  it 'should match the expected body', (done) ->
    extract = extractor.extract path.resolve(__dirname, "data/test3.doc")
    expect(extract).to.be.fulfilled
    extract
      .then (result) ->
        body = result.getBody()
        expect(body).to.match new RegExp('Can You Release Commercial Works\\?')
        expect(body).to.match new RegExp('Apache v2\\.0')
        expect(body).to.match new RegExp('you want your protections to be\\.')
        done()
