Buffer =         require('buffer').Buffer

oleDoc =         require('ole-doc').OleCompoundDoc
Promise =        require 'bluebird'

filters =        require './filters'
translations =   require './translations'

Document =       require './document'

class WordExtractor

  constructor: () ->


  ## Given an OLE stream, returns all the data in a buffer,
  ## as a promise.
  streamBuffer = (stream) ->
    new Promise (resolve, reject) ->
      chunks = []
      stream.on 'data', (chunk) ->
        chunks.push chunk
      stream.on 'error', (error) ->
        reject(error)
      stream.on 'end', () ->
        resolve(Buffer.concat(chunks))


  extractDocument = (filename) ->
    new Promise (resolve, reject) ->
      console.log "File", filename
      document = new oleDoc(filename)
      document.on 'err', (error) =>
        reject(error)
      document.on 'ready', () =>
        resolve(document)
      document.read()


  extract: (filename) ->
    extractDocument(filename)
      .then (document) ->
        documentStream(document, 'WordDocument')
          .then (stream) ->
            streamBuffer(stream)
          .then (buffer) ->
            extractWordDocument(document, buffer)


  documentStream = (document, stream) ->
    Promise.resolve document.stream(stream)


  writeBookmarks = (buffer, tableBuffer, result) ->
    fcSttbfBkmk = buffer.readUInt32LE(0x0142)
    lcbSttbfBkmk = buffer.readUInt32LE(0x0146)
    fcPlcfBkf = buffer.readUInt32LE(0x014a)
    lcbPlcfBkf = buffer.readUInt32LE(0x014e)
    fcPlcfBkl = buffer.readUInt32LE(0x0152)
    lcbPlcfBkl = buffer.readUInt32LE(0x0156)

    return if lcbSttbfBkmk == 0

    sttbfBkmk = tableBuffer.slice(fcSttbfBkmk, fcSttbfBkmk + lcbSttbfBkmk)
    plcfBkf = tableBuffer.slice(fcPlcfBkf, fcPlcfBkf + lcbPlcfBkf)
    plcfBkl = tableBuffer.slice(fcPlcfBkl, fcPlcfBkl + lcbPlcfBkl)

    fcExtend = sttbfBkmk.readUInt16LE(0)
    cData = sttbfBkmk.readUInt16LE(2)
    cbExtra = sttbfBkmk.readUInt16LE(4)

    if fcExtend != 0xffff
      throw new Error("Internal error: unexpected single-byte bookmark data")

    offset = 6
    index = 0
    bookmarks = {}

    while offset < lcbSttbfBkmk
      length = sttbfBkmk.readUInt16LE(offset)
      length = length * 2
      segment = sttbfBkmk.slice(offset + 2, offset + 2 + length)
      cpStart = plcfBkf.readUInt32LE(index * 4)
      cpEnd = plcfBkl.readUInt32LE(index * 4)
      result.bookmarks[segment] = {start: cpStart, end: cpEnd}
      offset = offset + length + 2


  writePieces = (buffer, tableBuffer, result) ->
    pos = buffer.readUInt32LE(0x01a2)
    console.log "Pos", pos

    while true
      flag = tableBuffer.readUInt8(pos)
      console.log "Flag skipping", flag
      break if flag != 1

      pos = pos + 1
      skip = tableBuffer.readUInt16LE(pos)
      pos = pos + 2 + skip
      console.log "Skipping", skip

    flag = tableBuffer.readUInt8(pos)
    console.log "Flag after skipping", flag
    pos = pos + 1
    if flag != 2
      throw new Error("Internal error: ccorrupted Word file")

    pieceTableSize = tableBuffer.readUInt32LE(pos)
    pos = pos + 4

    pieces = (pieceTableSize - 4) / 12
    start = 0

    for x in [0..pieces - 1] by 1
      offset = pos + ((pieces + 1) * 4) + (x * 8) + 2
      filePos = tableBuffer.readUInt32LE(offset)
      unicode = false
      if (filePos & 0x40000000) == 0
        unicode = true
      else
        filePos = filePos & ~(0x40000000)
        filePos = Math.floor(filePos / 2)
      lStart = tableBuffer.readUInt32LE(pos + (x * 4))
      lEnd = tableBuffer.readUInt32LE(pos + ((x + 1) * 4))
      totLength = lEnd - lStart

      console.log "lStart", lStart, "lEnd", lEnd

      piece = {
        start: start
        totLength: totLength
        filePos: filePos
        unicode: unicode
      }

      console.log "Piece", piece
      result.pieces.push getPiece(buffer, piece)

      start = start + (if unicode then Math.floor(totLength / 2) else totLength)


  extractWordDocument = (document, buffer) ->
    new Promise (resolve, reject) ->
      magic = buffer.readUInt16LE(0)
      if magic != 0xa5ec
        return reject new Error("This does not seem to be a Word document: Invalid magic number: " + magic)

      flags = buffer.readUInt16LE(0xA)
      console.log "Flags", flags.toString(16)

      table = if (flags & 0x0200) != 0 then "1Table" else "0Table"

      documentStream(document, table)
        .then (stream) ->
          streamBuffer stream

        .then (tableBuffer) ->
          result = new Document()
          result.boundaries.fcMin = buffer.readUInt32LE(0x0018)
          result.boundaries.ccpText = buffer.readUInt32LE(0x004c)
          result.boundaries.ccpFtn = buffer.readUInt32LE(0x0050)
          result.boundaries.ccpHdd = buffer.readUInt32LE(0x0054)
          result.boundaries.ccpAtn = buffer.readUInt32LE(0x005c)

          writeBookmarks buffer, tableBuffer, result
          writePieces buffer, tableBuffer, result

          resolve result

        .catch (error) ->
          reject error


  getPiece = (buffer, piece) ->
    pstart = piece.start
    ptotLength = piece.totLength
    pfilePos = piece.filePos
    punicode = piece.unicode

    pend = pstart + ptotLength
    textStart = pfilePos
    textEnd = textStart + (pend - pstart)

    if punicode
      piece.text = addUnicodeText buffer, textStart, textEnd
    else
      piece.text = addText buffer, textStart, textEnd


  addText = (buffer, textStart, textEnd) ->
    console.log "Adding text", textStart, textEnd
    slice = buffer.slice(textStart, textEnd)
    slice.toString('binary')

  addUnicodeText = (buffer, textStart, textEnd) ->
    slice = buffer.slice(textStart, 2*textEnd - textStart)
    string = slice.toString('ucs2')

    # See the conversion table for FcCompressed structures. Note that these
    # should not affect positions, as these are characters now, not bytes
    # for i in [0..string.length]
    #   if

    string


module.exports = WordExtractor
