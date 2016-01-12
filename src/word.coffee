Buffer =         require('buffer').Buffer

oleDoc =         require('ole-doc').OleCompoundDoc
Promise =        require 'bluebird'

filters =        require './filters'
translations =   require './translations'

class WordExtractor

  constructor: (options) ->
    if typeof options == 'string'
      options = {filename: options}
    else
      options ?= {}

    @filters = filters
    @fcTranslations = translations
    @filename = options.filename
    console.log "Filename", @filename


  streamToBuffer: (stream, cb) ->
    chunks = []
    stream.on 'data', (chunk) ->
      chunks.push chunk
    stream.on 'error', (error) ->
      cb error
    stream.on 'end', () ->
      cb null, Buffer.concat(chunks)


  extractBuffer: (buffer, cb) ->
    instance = @

    ## Check the magic number.
    magic = buffer.readUInt16LE(0)
    if magic != 0xa5ec
      return cb "Invalid magic number"

    flags = buffer.readUInt16LE(0xA)
    console.log "Flags", flags.toString(16)

    table = if (flags & 0x0200) != 0 then "1Table" else "0Table"

    stream = @doc.stream(table)
    @streamToBuffer stream, (error, data) =>
      return cb error if error?

      console.log "Read table buffer", data.length

      @tableData = data

      console.log "Table data", @tableData

      @fcMin = @data.readUInt32LE(0x0018)
      @ccpText = @data.readUInt32LE(0x004c)
      @ccpFtn = @data.readUInt32LE(0x0050)
      @ccpHdd = @data.readUInt32LE(0x0054)
      @ccpAtn = @data.readUInt32LE(0x005c)

      @charPLC = @data.readUInt32LE(0x00fa)
      @charPlcSize = @data.readUInt32LE(0x00fe)
      @parPLC = @data.readUInt32LE(0x0102)
      @parPlcSize = @data.readUInt32LE(0x0106)

      @fcPlcfBteChpx = @data.readUInt32LE(0x0fa)
      @lcbPlcfBteChpx = @data.readUInt32LE(0x0fe)

      console.log "fcMin", @fcMin
      console.log "ccpText", @ccpText

      @getBookmarks()
      @pieces = @getPieces()

      cb null


  getBookmarks: () ->
    @fcSttbfBkmk = @data.readUInt32LE(0x0142)
    @lcbSttbfBkmk = @data.readUInt32LE(0x0146)
    @fcPlcfBkf = @data.readUInt32LE(0x014a)
    @lcbPlcfBkf = @data.readUInt32LE(0x014e)
    @fcPlcfBkl = @data.readUInt32LE(0x0152)
    @lcbPlcfBkl = @data.readUInt32LE(0x0156)

    return if @lcbSttbfBkmk == 0

    @sttbfBkmk = @tableData.slice(@fcSttbfBkmk, @fcSttbfBkmk + @lcbSttbfBkmk)
    @plcfBkf = @tableData.slice(@fcPlcfBkf, @fcPlcfBkf + @lcbPlcfBkf)
    @plcfBkl = @tableData.slice(@fcPlcfBkl, @fcPlcfBkl + @lcbPlcfBkl)

    @fcExtend = @sttbfBkmk.readUInt16LE(0)
    @cData = @sttbfBkmk.readUInt16LE(2)
    @cbExtra = @sttbfBkmk.readUInt16LE(4)
    confess("Internal error: unexpected single-byte bookmark data") unless ($fcExtend == 0xffff);

    offset = 6
    index = 0
    bookmarks = {}

    while offset < @lcbSttbfBkmk
      length = @sttbfBkmk.readUInt16LE(offset)
      length = length * 2
      segment = @sttbfBkmk.slice(offset + 2, offset + 2 + length)
      cpStart = @plcfBkf.readUInt32LE(index * 4)
      cpEnd = @plcfBkl.readUInt32LE(index * 4)
      bookmarks[segment] = {start: cpStart, end: cpEnd}
      offset = offset + length + 2

    @bookmarks = bookmarks


  filter: (text, shouldFilter) ->

    return text if !shouldFilter

    replacer = (match, rest...) ->
      if match.length == 1
        replaced = @filters[match.charPointAt(0)]
        if replaced == 0
          ""
        else
          replaced
      else if rest.length == 2
        ""
      else if rest.length == 3
        rest[0]

    matcher = /(?:[\x02\x05\x07\x08\x0a\x0d\u2018\u2019\u201c\u201d\u2002\u2003\u2012\u2013\u2014]|\x13(?:[^\x14]*\x14)?([^\x15]*)\x15)/g
    text.replace replacer


  getTextRange: (start, end) ->
    pieces = @pieces
    startPiece = @getPieceIndex(start)
    endPiece = @getPieceIndex(end)
    result = []
    for i in [startPiece..endPiece] by 1
      piece = pieces[i]
      xstart = if i == startPiece then start - piece.position else 0
      xend = if i == endPiece then end - piece.position else piece.endPosition
      result.push piece.text.substring(xstart, xend - xstart)

    result.join("")


  getPieceIndex: (position) ->
    for piece, i in @pieces
      return i if position <= piece.endPosition


  getBody: (shouldFilter) ->
    shouldFilter ?= true
    string = @getTextRange(0, @ccpText)
    @filter string, shouldFilter


  getPieces: () ->
    pos = @data.readUInt32LE(0x01a2)
    console.log "Pos", pos
    result = []

    while true
      flag = @tableData.readUInt8(pos)
      console.log "Flag skipping", flag
      break if flag != 1

      pos = pos + 1
      skip = @tableData.readUInt16LE(pos)
      pos = pos + 2 + skip
      console.log "Skipping", skip

    flag = @tableData.readUInt8(pos)
    console.log "Flag after skipping", flag
    pos = pos + 1
    if flag != 2
      throw new Error("Internal error: ccorrupted Word file")

    pieceTableSize = @tableData.readUInt32LE(pos)
    pos = pos + 4

    pieces = (pieceTableSize - 4) / 12
    start = 0

    for x in [0..pieces - 1] by 1
      offset = pos + ((pieces + 1) * 4) + (x * 8) + 2
      filePos = @tableData.readUInt32LE(offset)
      unicode = false
      if (filePos & 0x40000000) == 0
        unicode = true
      else
        filePos = filePos & ~(0x40000000)
        filePos = Math.floor(filePos / 2)
      lStart = @tableData.readUInt32LE(pos + (x * 4))
      lEnd = @tableData.readUInt32LE(pos + ((x + 1) * 4))
      totLength = lEnd - lStart

      console.log "lStart", lStart, "lEnd", lEnd

      piece = {
        start: start
        totLength: totLength
        filePos: filePos
        unicode: unicode
      }

      console.log "Piece", piece
      result.push piece

      start = start + (if unicode then Math.floor(totLength / 2) else totLength)

    ## And return
    result


  getPiece: (piece) ->
    pstart = piece.start
    ptotLength = piece.totLength
    pfilePos = piece.filePos
    punicode = piece.unicode

    pend = pstart + ptotLength
    textStart = pfilePos
    textEnd = textStart + (pend - pstart)

    if punicode
      piece.text = @addUnicodeText textStart, textEnd
    else
      piece.text = @addText textStart, textEnd


  addText: (textStart, textEnd) ->
    console.log "Adding text", textStart, textEnd
    slice = @data.slice(textStart, textEnd)
    slice.toString('binary')

  addUnicodeText: (textStart, textEnd) ->
    slice = @data.slice(textStart, 2*textEnd - textStart)
    string = slice.toString('ucs2')

    # See the conversion table for FcCompressed structures. Note that these
    # should not affect positions, as these are characters now, not bytes
    # for i in [0..string.length]
    #   if

    string

  extract: (cb) ->
    console.log "File", @filename
    @doc = new oleDoc(@filename)
    @doc.on 'err', (err) =>
      console.log "Error", err
      cb err
    @doc.on 'ready', () =>
      stream = @doc.stream('WordDocument')
      @streamToBuffer stream, (error, data) =>
        return cb(error) if error?
        @data = data
        @extractBuffer data, (error) =>
          cb error
    @doc.read()


module.exports = WordExtractor
