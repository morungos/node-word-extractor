chai = require('chai')
chai.use(require('chai-as-promised'))
should = chai.should()

Word = require('../lib/word')

describe 'Word', () ->

  describe 'test1.doc', () ->

    it 'should extract successfully', (done) ->

      word = new Word('test/data/test1.doc')
      word.extract (err, data) ->
        console.log "Error", err, data
        done(err)
