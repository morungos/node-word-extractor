chai = require('chai')
expect = chai.expect

fs = require('fs')
path = require('path')
Buffer = require('buffer').Buffer

WordExtractor = require('../src/word')
Document = require('../src/document')

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
      .catch (err) -> done(err)

  it 'should retrieve the annotations', (done) ->
    extract = extractor.extract path.resolve(__dirname, "data/test10.doc")
    expect(extract).to.be.fulfilled
    extract
      .then (result) ->
        annotations = result.getAnnotations()
        expect(annotations).to.match new RegExp('Second paragraph comment')
        expect(annotations).to.match new RegExp('Third paragraph comment')
        expect(annotations).to.match new RegExp('this is all I have to say on the matter')
        done()
      .catch (err) -> done(err)
