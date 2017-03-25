chai = require('chai')
expect = chai.expect

fs = require('fs')
path = require('path')
Buffer = require('buffer').Buffer

WordExtractor = require('../src/word')
Document = require('../src/document')

describe 'Word file test12.doc', () ->

  extractor = new WordExtractor()

  it 'should match the expected body', () ->
    extractor.extract path.resolve(__dirname, "data/test12.doc")
      .then (result) ->
        body = result.getBody()
        expect(body).to.match /Row 2, cell 3/
