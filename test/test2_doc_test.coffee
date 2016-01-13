chai = require('chai')
chai.use(require('chai-as-promised'))
should = chai.should()

WordExtractor = require('../lib/word')
Document = require('../lib/document')

describe 'test2.doc', () ->

  extractor = new WordExtractor()

  it 'should extract a document successfully', (done) ->
    extractor.extract 'test/data/test2.doc'
      .should.be.fulfilled
        .then (result) ->
          result.should.be.an.instanceof(Document)
          result.pieces.should.be.instanceof(Array)
          result.boundaries.should.contain.keys(['fcMin', 'ccpText', 'ccpFtn', 'ccpHdd', 'ccpAtn'])
        .should.notify(done)

  # it 'should retrieve document text', (done) ->
  #   extractor.extract 'test/data/test2.doc'
  #     .then (document) ->
  #       body = document.getBody()
  #       body.should.match new RegExp('My name is Ryan')
  #       body.should.match new RegExp('create several FKPs for testing purposes')
  #       body.should.match new RegExp('dsadasdasdasd')
  #     .should.notify(done)
