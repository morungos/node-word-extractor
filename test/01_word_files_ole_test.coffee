chai = require('chai')
expect = chai.expect

fs = require('fs')
path = require('path')
Buffer = require('buffer').Buffer

oleDoc = require('../src/ole-doc').OleCompoundDoc;

testWordFile = (file) ->
  describe 'Word file ' + file, () ->

    it 'can be opened correctly', (done) ->
      filename = path.resolve(__dirname, "data/" + file)
      doc = new oleDoc(filename)
      doc.on 'err', -> done(err)
      doc.on 'ready', -> done()
      doc.read()

    it 'generates a valid Word stream', (done) ->
      filename = path.resolve(__dirname, "data/" + file)
      doc = new oleDoc(filename)
      doc.on 'err', (err) -> done(err)
      doc.on 'ready', () ->
        chunks = []
        stream = doc.stream('WordDocument')
        stream.on 'data', (chunk) -> chunks.push(chunk)
        stream.on 'error', (error) -> done(error)
        stream.on 'end', () ->
          buffer = Buffer.concat(chunks)
          magicNumber = buffer.readUInt16LE(0)
          expect(magicNumber.toString(16)).to.equal("a5ec")
          done()
      doc.read()

files = fs.readdirSync(path.resolve(__dirname, "data"))
files.filter (file) ->
  testWordFile(file) if /\.doc$/i.test(file)
