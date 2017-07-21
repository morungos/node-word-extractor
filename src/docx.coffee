xpath = require('xpath')
Dom = require('xmldom').DOMParser
yauzl = require('yauzl')

getTextFromZipFile = (zipfile, entry) ->
  new Promise (resolve, reject) ->
    zipfile.openReadStream entry, (err, readStream) ->

      return reject(err) if err

      text = ''
      error = ''

      readStream.on 'data', (chunk) -> text += chunk

      readStream.on 'end', () ->
        return reject(error) if error.length > 0
        resolve(text)

      readStream.on 'error', (err) -> error += err


calculateExtractedText = (inText) ->
  doc = new Dom().parseFromString(inText)
  ps = xpath.select("//*[local-name()='p']", doc)
  text = ''

  ps.forEach (paragraph) ->
    localText = ''

    parent = paragraph.parentNode
    if parent.localName == 'tc'
      console.log parent.parentNode.localName

    paragraph = new Dom().parseFromString(paragraph.toString())

    ts = xpath.select("//*[local-name()='t' or local-name()='tab' or local-name()='br' or local-name()='instrText']", paragraph)
    ts.forEach (t) ->
      if t.localName == 't' && t.childNodes.length > 0
        localText += t.childNodes[0].data
      else if t.localName == 'tab' or t.localName == 'br'
        localText += ' '
      else if t.localName == 'instrText'
        localText += t.childNodes[0].data

    text += localText + '\n'

  return text


class DOCXExtractor

  constructor: () ->

  extract: (doc) ->
    new Promise (resolve, reject) ->
      yauzl.open doc, ( err, zipfile ) ->
        processedEntries = 0
        result = ''

        return reject(err) if err

        processEnd = () ->
          if zipfile.entryCount == ++processedEntries
            if result.length
              resolve(result)
            else
              reject(new Error("Extraction could not find content in file"))

        zipfile.on 'entry', (entry) ->
          if /\.xml$/.test(entry.fileName) and (! /^(word\/media\/|word\/_rels\/)/.test(entry.fileName))
            getTextFromZipFile(zipfile, entry)
              .then (data) ->
                text = calculateExtractedText(data)
                result += text + '\n'
                processEnd()
          else
            processEnd()

        zipfile.on 'error', (error) ->
          reject(err)


module.exports = DOCXExtractor;
