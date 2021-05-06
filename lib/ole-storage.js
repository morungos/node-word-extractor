var es = require('event-stream');

function Storage( doc, dirEntry ) {
  this._doc = doc;
  this._dirEntry = dirEntry;
}

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
    }

    if ( shortStream ) {
      doc._readShortSector( secIds[i], sectorCallback );
    }
    else {
      doc._readSector( secIds[i], sectorCallback );
    }
  });
};

module.exports = Storage;