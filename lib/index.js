const fs = require('fs');

class WordExtractor {

  constructor() {}

  getFirstBlock(doc) {
    return new Promise(function(resolve, reject) {
      return fs.open(doc, 'r', function(err, fd) {
        if (err) { 
          return reject(err); 
        }
        const buffer = Buffer.alloc(512);
        return fs.read(fd, buffer, 0, 512, 0, (err, read, buffer) =>
          fs.close(fd, function(err2) {
            if (err2) { 
              return reject(err2); 
            }
            return resolve(buffer.slice(0, read));
          })
        );
      });
    });
  }


  getFileType(doc) {
    return this.getFirstBlock(doc)
      .then(function(buffer) {
        if (buffer.readUInt16BE(0) === 0x504b) {
          const next = buffer.readUInt16BE(2);
          if ((next === 0x0304) || (next === 0x0506) || (next === 0x0708)) {
            return 'DOCX';
          }
        } else if (buffer.readUInt16BE(0) === 0xd0cf) {
          return 'DOC';
        } else {
          return null;
        }
      });
  }


  extract(doc) {
    return this.getFileType(doc);
  }
}


module.exports = WordExtractor;
