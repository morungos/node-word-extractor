
/**
 * Implements the main Open Office format extractor. Open Office .docx files
 * are essentially zip files containing streams, and each of these streams contains
 * XML content in one form or another. So we need to use {@link zlib} to extract
 * the streams, and something like `sax-js` to parse the XML that we find 
 * there. 
 * 
 * We probably don't need the whole of the Open Office data, we're only likely
 * to need a few streams. Sadly, the documentation for the file format is literally
 * 5000 pages.
 */

const SAX = require("sax");
const yauzl = require('yauzl');

const BufferReader = require('./buffer-reader');
const FileReader = require('./file-reader');
const OpenOfficeDocument = require('./open-office-document');

class OpenOfficeExtractor {

  constructor() { 
    this._streamTypes = {
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml': true,
      'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml': true,
      'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml': true
    };
    this._actions = {};
    this._document = new OpenOfficeDocument();
  }

  extract(reader) {
    //throw new Error("Extracting from .docx files not yet implemented");

    let archive = null;

    if (BufferReader.isBufferReader(reader)) {
      archive = Promise.resolve(yauzl.fromBuffer(reader.buffer()));
    } else if (FileReader.isFileReader(reader)) {
      archive = new Promise((resolve, reject) => {
        yauzl.fromFd(reader.fd(), {lazyEntries: true, autoClose: false}, function(err, zipfile) {
          if (err) {
            return reject(err);
          }
          resolve(zipfile);
        });
      });
    }

    return archive
      .then((zipfile) => {
        return new Promise((resolve, reject) => {
          zipfile.readEntry();
          zipfile.on("error", function(error) {
            reject(error);
          });
          zipfile.on("entry", (entry) => {
            if ('[Content_Types].xml' === entry.fileName || this._actions[entry.fileName]) {
              //console.log("entry", entry.fileName);
              return this.handleEntry(zipfile, entry)
                .then(() => {
                  zipfile.readEntry();
                });
            }
  
            zipfile.readEntry();
          });
          zipfile.on("end", () => {
            //console.log(this._document);
            resolve(this._document);
          });  
        });
      });

  }

  createXmlParser() {
    const strict = true;
    const parser = SAX.createStream(strict); 

    parser.on("opentag", (node) => {

      if (node.name === 'Override') {
        const actionFunction = this._streamTypes[node.attributes['ContentType']];
        if (actionFunction) {
          const partName = node.attributes['PartName'].replace(/^[/]+/, '');
          const action = {action: actionFunction, type: node.attributes['ContentType']};
          // console.log("Marking", partName, action);
          this._actions[partName] = action;
        }
      } else if (node.name === 'w:document') {
        this._inDocument = true;
      }
      
    });

    parser.on('closetag', (node) => {
      if (node === 'w:document') {
        this._inDocument = false;
      } else if (node === 'w:p' && this._inDocument) {
        this._document.add("\n");
      }
    });

    parser.on('text', (string) => {
      if (this._inDocument) {
        this._document.add(string);
      }
    });

    return parser;
  }

  handleEntry(zipfile, entry) {
    return new Promise((resolve, reject) => {
      zipfile.openReadStream(entry, (err, readStream) => {
        if (err) {
          return reject(err);
        }

        const parser = this.createXmlParser();
        readStream.on("end", resolve);
        readStream.pipe(parser);
      });  
    });
  }

}

module.exports = OpenOfficeExtractor;
