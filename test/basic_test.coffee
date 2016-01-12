chai = require('chai')
chai.use(require('chai-as-promised'))
should = chai.should()

WordExtractor = require('../lib/word')

describe 'WordExtractor', () ->

  describe 'test1.doc', () ->

    it 'should extract successfully', (done) ->

      word = new WordExtractor()
      word.extract 'test/data/test1.doc', (err, data) ->
        console.log "Error", err, data
        done(err)
