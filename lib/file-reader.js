const fs = require('fs');

class FileReader {

  constructor(filename) {
    this._filename = filename;
  }

  open() {
    return new Promise((resolve, reject) => {
      fs.open(this._filename, 'r', 0o666, (err, fd) => {
        if(err) {
          return reject(err);
        }
  
        this._fd = fd;
        resolve();
      });  
    });
  }

  /**
   * Closes the file descriptor associated with an open document, if there
   * is one, and returns a promise that resolves when the file handle is closed.
   * @returns a promise
   */
  close() {
    return new Promise((resolve, reject) => {
      if (this._fd) {
        fs.close(this._fd, (err) => {
          if (err) {
            return reject(err);
          }
          delete this._fd;
          resolve();
        });
      } else {
        resolve();
      }
    });
  }


  read(buffer, offset, length, position) {
    return new Promise((resolve, reject) => {
      if ( !this._fd) {
        return reject(new Error("file not open"));
      }
      fs.read(this._fd, buffer, offset, length, position, (err, bytesRead, buffer) => {
        if (err) {
          return reject(err);
        }
        resolve(buffer);
      });
    });
  }
}

module.exports = FileReader;
