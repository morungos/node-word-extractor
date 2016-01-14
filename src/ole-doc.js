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
var util = require('util');
var async = require('async');
var _ = require('underscore');
var es = require('event-stream');

function Header() {
};

Header.ole_id = new Buffer( 'D0CF11E0A1B11AE1', 'hex' );

Header.prototype.load = function( buffer ) {
   var i;
   for( i = 0; i < 8; i++ ) {
      if ( Header.ole_id[i] != buffer[i] )
         return false;
   }

   this.secSize        = 1 << buffer.readInt16LE( 30 );  // Size of sectors
   this.shortSecSize   = 1 << buffer.readInt16LE( 32 );  // Size of short sectors
   this.SATSize        =      buffer.readInt32LE( 44 );  // Number of sectors used for the Sector Allocation Table
   this.dirSecId       =      buffer.readInt32LE( 48 );  // Starting Sec ID of the directory stream
   this.shortStreamMax =      buffer.readInt32LE( 56 );  // Maximum size of a short stream
   this.SSATSecId      =      buffer.readInt32LE( 60 );  // Starting Sec ID of the Short Sector Allocation Table
   this.SSATSize       =      buffer.readInt32LE( 64 );  // Number of sectors used for the Short Sector Allocation Table
   this.MSATSecId      =      buffer.readInt32LE( 68 );  // Starting Sec ID of the Master Sector Allocation Table
   this.MSATSize       =      buffer.readInt32LE( 72 );  // Number of sectors used for the Master Sector Allocation Table

   // The first 109 sectors of the MSAT
   this.partialMSAT = new Array(109);
   for( i = 0; i < 109; i++ )
      this.partialMSAT[i] = buffer.readInt32LE( 76 + i * 4 );

   return true;
};

function AllocationTable(doc) {
   this._doc = doc;
}

AllocationTable.SecIdFree       = -1;
AllocationTable.SecIdEndOfChain = -2;
AllocationTable.SecIdSAT        = -3;
AllocationTable.SecIdMSAT       = -4;

AllocationTable.prototype.load = function(secIds, callback) {
   var self = this;
   var doc = self._doc;
   var header = doc._header;

   self._table = new Array( secIds.length * ( header.secSize / 4 ) );

   doc._readSectors( secIds, function(buffer) {
      var i;
      for ( i = 0; i < buffer.length / 4; i++ ) {
         self._table[i] = buffer.readInt32LE( i * 4 );
      }
      callback();
   });
};

AllocationTable.prototype.getSecIdChain = function(startSecId) {
   var secId = startSecId;
   var secIds = [];
   while ( secId != AllocationTable.SecIdEndOfChain ) {
      secIds.push( secId );
      secId = this._table[secId];
   }

   return secIds;
};

function DirectoryTree(doc) {
   this._doc = doc;
}

DirectoryTree.EntryTypeEmpty   = 0;
DirectoryTree.EntryTypeStorage = 1;
DirectoryTree.EntryTypeStream  = 2;
DirectoryTree.EntryTypeRoot    = 5;

DirectoryTree.NodeColorRed   = 0;
DirectoryTree.NodeColorBlack = 1;

DirectoryTree.Leaf = -1;

DirectoryTree.prototype.load = function( secIds, callback ) {
   var self = this;
   var doc = this._doc;

   doc._readSectors( secIds, function(buffer) {

      var count = buffer.length / 128;
      self._entries = new Array( count );
      var i = 0;
      for( i = 0; i < count; i++ )
      {
         var offset = i * 128;

         var nameLength = Math.max( buffer.readInt16LE( 64 + offset ) - 1, 0 );

         var entry = {};
         entry.name = buffer.toString('utf16le', 0 + offset, nameLength + offset );
         entry.type = buffer.readInt8( 66 + offset );
         entry.nodeColor = buffer.readInt8( 67 + offset );
         entry.left = buffer.readInt32LE( 68 + offset );
         entry.right = buffer.readInt32LE( 72 + offset );
         entry.storageDirId = buffer.readInt32LE( 76 + offset );
         entry.secId = buffer.readInt32LE( 116 + offset );
         entry.size = buffer.readInt32LE( 120 + offset );

         self._entries[i] = entry;
      }

      self.root = _.find( self._entries, function(entry) {
         return entry.type === DirectoryTree.EntryTypeRoot;
      });

      self._buildHierarchy( self.root );

      callback();
   });
};

DirectoryTree.prototype._buildHierarchy = function( storageEntry ) {
   var self = this;
   var childIds = this._getChildIds( storageEntry );

   storageEntry.storages = {};
   storageEntry.streams  = {};

   _.each( childIds, function( childId ) {
      var childEntry = self._entries[childId];
      var name = childEntry.name;
      if ( childEntry.type === DirectoryTree.EntryTypeStorage ) {
         storageEntry.storages[name] = childEntry;
      }
      if ( childEntry.type === DirectoryTree.EntryTypeStream ) {
         storageEntry.streams[name] = childEntry;
      }
   });

   _.each( storageEntry.storages, function( childStorageEntry ) {
      self._buildHierarchy( childStorageEntry );
   });
};

DirectoryTree.prototype._getChildIds = function( storageEntry ) {
   var self = this;
   var childIds = [];

   function visit( visitEntry ) {
      if ( visitEntry.left !== DirectoryTree.Leaf ) {
         childIds.push( visitEntry.left );
         visit( self._entries[visitEntry.left] );
      }
      if ( visitEntry.right !== DirectoryTree.Leaf ) {
         childIds.push( visitEntry.right );
         visit( self._entries[visitEntry.right] );
      }
   };

   if ( storageEntry.storageDirId > -1 ) {
      childIds.push( storageEntry.storageDirId );
      var rootChildEntry = self._entries[storageEntry.storageDirId];
      visit( rootChildEntry );
   }

   return childIds;
};

function Storage( doc, dirEntry ) {
   this._doc = doc;
   this._dirEntry = dirEntry;
};

Storage.prototype.storage = function( storageName ) {
   return new Storage( this._doc, this._dirEntry.storages[ storageName ] );
};

Storage.prototype.stream = function( streamName ) {
   var streamEntry = this._dirEntry.streams[streamName];
   if ( !streamEntry )
      return null;

   var self = this;
   var doc  = self._doc;
   var bytes = streamEntry.size;

   var allocationTable = doc._SAT;
   var shortStream = false;
   if ( bytes < doc._header.shortStreamMax ) {
      shortStream = true;
      allocationTable = doc._SSAT;
   }

   var secIds = allocationTable.getSecIdChain( streamEntry.secId );

   return es.readable( function( i, callback ) {
      var stream = this;  // Function called in context of stream

      if ( i >= secIds.length ) {
         stream.emit('end');
         return;
      }

      function sectorCallback(buffer) {
         if ( bytes - buffer.length < 0 ) {
            buffer = buffer.slice( 0, bytes );
         }

         bytes -= buffer.length;
         stream.emit('data', buffer);
         callback();
      };

      if ( shortStream ) {
         doc._readShortSector( secIds[i], sectorCallback );
      }
      else {
         doc._readSector( secIds[i], sectorCallback );
      }
   });
};

//function Stream( doc, dirEntry ) {
//   this._doc = doc;
//   this._dirEntry = dirEntry;
//};

function OleCompoundDoc( filename ) {
   EventEmitter.call(this);

   this._filename = filename;
   this._skipBytes = 0;
};
util.inherits(OleCompoundDoc, EventEmitter);

OleCompoundDoc.prototype.read = function() {
   this._read();
};

OleCompoundDoc.prototype.readWithCustomHeader = function( size, callback ) {
   this._skipBytes = size;
   this._customHeaderCallback = callback;
   this._read();
};

OleCompoundDoc.prototype._read = function() {
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
};

OleCompoundDoc.prototype._openFile = function( callback ) {
   var self = this;
   fs.open( this._filename, 'r', 0666, function(err, fd) {
      if( err ) {
         self.emit('err', err);
         return;
      }

      self._fd = fd;
      callback();
   });
};

OleCompoundDoc.prototype._readCustomHeader = function(callback) {
   var self = this;
   var buffer = new Buffer(this._skipBytes);
   fs.read( self._fd, buffer, 0, this._skipBytes, 0, function(err, bytesRead, buffer) {
      if(err) {
         self.emit('err', err);
         return;
      }

      if( !self._customHeaderCallback(buffer) )
         return;

      callback();
   });
};

OleCompoundDoc.prototype._readHeader = function(callback) {
   var self = this;
   var buffer = new Buffer(512);
   fs.read( this._fd, buffer, 0, 512, 0 + this._skipBytes, function(err, bytesRead, buffer) {
      if( err ) {
         self.emit('err', err);
         return;
      }

      var header = self._header = new Header();
      if ( !header.load( buffer ) ) {
         self.emit('err', 'Not a valid compound document.');
         return;
      }

      callback();
   });
};

OleCompoundDoc.prototype._readMSAT = function(callback) {
   var self = this;
   var header = self._header;

   self._MSAT = header.partialMSAT.slice(0);
   self._MSAT.length = header.SATSize;

   if( header.SATSize <= 109 || header.MSATSize == 0 ) {
      callback();
      return;
   }

   var buffer = new Buffer( header.secSize );
   var currMSATIndex = 109;
   var i = 0;
   var secId = header.MSATSecId;

   async.whilst(
      function() {
         return i < header.MSATSize;
      },
      function(whilstCallback) {
         self._readSector(secId, function(sectorBuffer) {
            var s;
            for( s = 0; s < header.secSize - 4; s += 4 )
            {
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
};

OleCompoundDoc.prototype._readSector = function(secId, callback) {
   this._readSectors( [ secId ], callback );
};

OleCompoundDoc.prototype._readSectors = function(secIds, callback) {
   var self = this;
   var header = self._header;
   var buffer = new Buffer( secIds.length * header.secSize );

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
};

OleCompoundDoc.prototype._readShortSector = function(secId, callback) {
   this._readShortSectors( [ secId ], callback );
};

OleCompoundDoc.prototype._readShortSectors = function(secIds, callback) {
   var self = this;
   var header = self._header;
   var buffer = new Buffer( secIds.length * header.shortSecSize );

   var i = 0;

   async.whilst(
      function() {
         return i < secIds.length;
      },
      function(whilstCallback) {
         var bufferOffset = i * header.shortSecSize;
         var fileOffset = self._getFileOffsetForShortSec( secIds[i] );
         fs.read( self._fd, buffer, bufferOffset, header.shortSecSize, fileOffset, function(err, bytesRead, buffer) {
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
};

OleCompoundDoc.prototype._readSAT = function(callback) {
   var self = this;
   self._SAT = new AllocationTable(self);

   self._SAT.load( self._MSAT, callback );
};

OleCompoundDoc.prototype._readSSAT = function(callback) {
   var self = this;
   var header = self._header;
   self._SSAT = new AllocationTable(self);

   var secIds = self._SAT.getSecIdChain( header.SSATSecId );
   if ( secIds.length != header.SSATSize ) {
      self.emit('err', 'Invalid Short Sector Allocation Table');
      return;
   }

   self._SSAT.load( secIds, callback);
};

OleCompoundDoc.prototype._readDirectoryTree = function(callback) {
   var self = this;
   var header = self._header;

   self._directoryTree = new DirectoryTree(this);

   var secIds = self._SAT.getSecIdChain( header.dirSecId );
   self._directoryTree.load( secIds, function() {

      var rootEntry = self._directoryTree.root;
      self._rootStorage = new Storage( self, rootEntry );
      self._shortStreamSecIds = self._SAT.getSecIdChain( rootEntry.secId );

      callback();
   });
};

OleCompoundDoc.prototype._getFileOffsetForSec = function( secId ) {
   var secSize = this._header.secSize;
   return this._skipBytes + (secId + 1) * secSize;  // Skip past the header sector
};

OleCompoundDoc.prototype._getFileOffsetForShortSec = function( shortSecId ) {
   var shortSecSize = this._header.shortSecSize;
   var shortStreamOffset = shortSecId * shortSecSize;

   var secSize = this._header.secSize;
   var secIdIndex = Math.floor( shortStreamOffset / secSize );
   var secOffset = shortStreamOffset % secSize;
   var secId = this._shortStreamSecIds[secIdIndex];

   return this._getFileOffsetForSec( secId ) + secOffset;
};

OleCompoundDoc.prototype.storage = function( storageName ) {
   return this._rootStorage.storage( storageName );
};

OleCompoundDoc.prototype.stream = function( streamName ) {
   return this._rootStorage.stream( streamName );
};

exports.OleCompoundDoc = OleCompoundDoc;
