filters =        require './filters'

class Document


  constructor: () ->
    @pieces = []
    @bookmarks = {}
    @boundaries = {}


  getPieceIndex = (pieces, position) ->
    for piece, i in pieces
      return i if position <= piece.endPosition


  getTextRange: (start, end) ->
    pieces = @pieces
    startPiece = getPieceIndex(pieces, start)
    endPiece = getPieceIndex(pieces, end)
    result = []
    for i in [startPiece..endPiece] by 1
      piece = pieces[i]
      xstart = if i == startPiece then start - piece.position else 0
      xend = if i == endPiece then end - piece.position else piece.endPosition
      result.push piece.text.substring(xstart, xend)

    result.join("")


  filter = (text, shouldFilter) ->

    return text if !shouldFilter

    replacer = (match, rest...) ->
      if match.length == 1
        replaced = filters[match.charCodeAt(0)]
        if replaced == 0
          ""
        else
          replaced
      else if rest.length == 2
        ""
      else if rest.length == 3
        rest[0]

    matcher = /(?:[\x02\x05\x07\x08\x0a\x0d\u2018\u2019\u201c\u201d\u2002\u2003\u2012\u2013\u2014]|\x13(?:[^\x14]*\x14)?([^\x15]*)\x15)/g
    text.replace(matcher, replacer)


  getBody: (shouldFilter) ->
    shouldFilter ?= true
    start = 0
    string = @getTextRange(start, start + @boundaries.ccpText)
    filter string, shouldFilter


  getFootnotes: (shouldFilter) ->
    shouldFilter ?= true
    start = @boundaries.ccpText
    string = @getTextRange(start, start + @boundaries.ccpFtn)
    filter string, shouldFilter


  getHeaders: (shouldFilter) ->
    shouldFilter ?= true
    start = @boundaries.ccpText + @boundaries.ccpFtn
    string = @getTextRange(start, start + @boundaries.ccpHdd)
    filter string, shouldFilter


  getAnnotations: (shouldFilter) ->
    shouldFilter ?= true
    start = @boundaries.ccpText + @boundaries.ccpFtn + @boundaries.ccpHdd
    string = @getTextRange(start, start + @boundaries.ccpAtn)
    filter string, shouldFilter


module.exports = Document
