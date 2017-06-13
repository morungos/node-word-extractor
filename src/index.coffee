fs = require('fs')


class WordExtractor

  constructor: () ->

  getFirstBlock: (doc) ->
    new Promise (resolve, reject) ->
      fs.open doc, 'r', (err, fd) ->
        return reject(err) if err
        buffer = Buffer.alloc(512)
        fs.read fd, buffer, 0, 512, 0, (err, read, buffer) ->
          fs.close fd, (err2) ->
            return reject(err) if err
            resolve buffer.slice(0, read)


  getFileType: (doc) ->
    this.getFirstBlock doc
      .then (buffer) ->
        if buffer.readUInt16BE(0) == 0x504b
          next = buffer.readUInt16BE(2)
          if next == 0x0304 or next == 0x0506 or next == 0x0708
            return 'DOCX'
        else if buffer.readUInt16BE(0) == 0xd0cf
          return 'DOC'
        else
          return null


  extract: (doc) ->
    this.getFileType doc
      .then (result) ->
        console.log result


module.exports = WordExtractor
