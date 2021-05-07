/* eslint-disable */
// Copyright (c) 2012 Chris Geiersbach
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//
// This component as adapted from node-ole-doc, available at:
// https://github.com/atariman486/node-ole-doc.
//
// WARNING: This embedded component will be removed in a future
// release. It is only included as there are some fixes which
// are not yet pushed into the npm distribution of node-ole-doc.

var fs = require('fs');
var EventEmitter = require('events').EventEmitter;
var async = require('async');

const Header = require('./ole-header');
const AllocationTable = require('./ole-allocation-table');
const DirectoryTree = require('./ole-directory-tree');
const Storage = require('./ole-storage');

class OleCompoundDoc extends EventEmitter {

  constructor(filename) {
    super();
    
    this._filename = filename;
    this._skipBytes = 0;
  }

  read() {
    this._read();
  }

  readWithCustomHeader( size, callback ) {
    this._skipBytes = size;
    this._customHeaderCallback = callback;
    this._read();
  }

  _read() {
    var series = [
      this._openFile.bind(this),
      this._readHeader.bind(this),
      this._readMSAT.bind(this),
      this._readSAT.bind(this),
      this._readSSAT.bind(this),
      this._readDirectoryTree.bind(this)
    ];

    if ( this._skipBytes != 0 ) {
      series.splice( 1, 0, this._readCustomHeader.bind(this) );
    }

    async.series(
      series,
      (function(err) {
         if ( err ) {
            this.emit('err', err);
            return;
         }

         this.emit('ready');
      }).bind(this)
    );
  }

  _openFile( callback ) {
    fs.open( this._filename, 'r', 0o666, (err, fd) => {
      if( err ) {
         this.emit('err', err);
         return;
      }

      this._fd = fd;
      callback();
    });
  }

  _readCustomHeader(callback) {
    const buffer = Buffer.alloc(this._skipBytes);
    fs.read( this._fd, buffer, 0, this._skipBytes, 0, (err, bytesRead, buffer) => {
      if(err) {
         this.emit('err', err);
         return;
      }

      if( !this._customHeaderCallback(buffer) )
         return;

      callback();
    });
  }

  _readHeader(callback) {
    const buffer = Buffer.alloc(512);
    fs.read( this._fd, buffer, 0, 512, 0 + this._skipBytes, (err, bytesRead, buffer) => {
      if( err ) {
        this.emit('err', err);
        return;
      }

      const header = this._header = new Header();
      if ( !header.load( buffer ) ) {
        this.emit('err', new Error("Not a valid compound document"));
        return;
      }

      callback();
    });
  }

  _readMSAT(callback) {
    var self = this;
    const header = self._header;

    self._MSAT = header.partialMSAT.slice(0);
    self._MSAT.length = header.SATSize;

    if( header.SATSize <= 109 || header.MSATSize == 0 ) {
      callback();
      return;
    }

    let currMSATIndex = 109;
    let i = 0;
    let secId = header.MSATSecId;

    async.whilst(
      function() {
        return i < header.MSATSize;
      },
      function(whilstCallback) {
        self._readSector(secId, function(sectorBuffer) {
          var s;
          for( s = 0; s < header.secSize - 4; s += 4 ) {
            if( currMSATIndex >= header.SATSize )
              break;
            else
              self._MSAT[currMSATIndex] = sectorBuffer.readInt32LE( s );

            currMSATIndex++;
          }

          secId = sectorBuffer.readInt32LE( header.secSize - 4 );
          i++;
          whilstCallback();
        });
      },
      function(err) {
        if ( err ) {
          self.emit('err', err);
          return;
        }

        callback();
      }
    );
  }

  _readSector(secId, callback) {
    this._readSectors( [ secId ], callback );
  };

  _readSectors(secIds, callback) {
    var self = this;
    var header = self._header;
    var buffer = Buffer.alloc( secIds.length * header.secSize );

    var i = 0;

    async.whilst(
      function() {
        return i < secIds.length;
      },
      function(whilstCallback) {
        var bufferOffset = i * header.secSize;
        var fileOffset = self._getFileOffsetForSec( secIds[i] );
        fs.read( self._fd, buffer, bufferOffset, header.secSize, fileOffset, function(err, bytesRead, buffer) {
          if ( err ) {
            self.emit('err', err);
            return;
          }
          i++;
          whilstCallback();
        });
      },
      function(err) {
        if ( err ) {
          self.emit('err', err);
        }
        callback(buffer);
      }
    );
  }

  _readShortSector(secId, callback) {
    this._readShortSectors( [ secId ], callback );
  }

  _readShortSectors(secIds, callback) {
    var self = this;
    const header = self._header;
    const buffer = Buffer.alloc( secIds.length * header.shortSecSize );

    let i = 0;

    async.whilst(
      function() {
        return i < secIds.length;
      },
      function(whilstCallback) {
        const bufferOffset = i * header.shortSecSize;
        var fileOffset = self._getFileOffsetForShortSec( secIds[i] );
        fs.read( self._fd, buffer, bufferOffset, header.shortSecSize, fileOffset, (err, bytesRead, buffer) => {
          if ( err ) {
            self.emit('err', err);
            return;
          }
          i++;
          whilstCallback();
        });
      },
      function(err) {
        if ( err ) {
          self.emit('err', err);
        }
        callback(buffer);
      }
    );
  }

  _readSAT(callback) {
    this._SAT = new AllocationTable(this);
    this._SAT.load( this._MSAT, callback );
  }

  _readSSAT(callback) {
    const header = this._header;
    this._SSAT = new AllocationTable(this);

    var secIds = this._SAT.getSecIdChain( header.SSATSecId );
    if ( secIds.length != header.SSATSize ) {
      this.emit('err', new Error("Invalid Short Sector Allocation Table"));
      return;
    }

    this._SSAT.load( secIds, callback);
  }

  _readDirectoryTree(callback) {
    const header = this._header;

    this._directoryTree = new DirectoryTree(this);

    const secIds = this._SAT.getSecIdChain( header.dirSecId );
    this._directoryTree.load( secIds, () => {

      const rootEntry = this._directoryTree.root;
      this._rootStorage = new Storage( this, rootEntry );
      this._shortStreamSecIds = this._SAT.getSecIdChain( rootEntry.secId );

      callback();
    });
  }

  _getFileOffsetForSec( secId ) {
    const secSize = this._header.secSize;
    return this._skipBytes + (secId + 1) * secSize;  // Skip past the header sector
  }

  _getFileOffsetForShortSec( shortSecId ) {
    const shortSecSize = this._header.shortSecSize;
    const shortStreamOffset = shortSecId * shortSecSize;

    const secSize = this._header.secSize;
    const secIdIndex = Math.floor( shortStreamOffset / secSize );
    const secOffset = shortStreamOffset % secSize;
    const secId = this._shortStreamSecIds[secIdIndex];

    return this._getFileOffsetForSec( secId ) + secOffset;
  }

  storage( storageName ) {
    return this._rootStorage.storage( storageName );
  }

  stream( streamName ) {
    return this._rootStorage.stream( streamName );
  }

}

module.exports = OleCompoundDoc;
