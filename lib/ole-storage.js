var es = require('event-stream');

class Storage {

  constructor(doc, dirEntry) {
    this._doc = doc;
    this._dirEntry = dirEntry;
  }

  storage( storageName ) {
    return new Storage( this._doc, this._dirEntry.storages[ storageName ] );
  }
  
  stream( streamName ) {
    const streamEntry = this._dirEntry.streams[streamName];
    if ( !streamEntry )
      return null;
  
    const doc  = this._doc;
    let bytes = streamEntry.size;
  
    let allocationTable = doc._SAT;
    let shortStream = false;
    if ( bytes < doc._header.shortStreamMax ) {
      shortStream = true;
      allocationTable = doc._SSAT;
    }
  
    const secIds = allocationTable.getSecIdChain( streamEntry.secId );
  
    return es.readable( function( i, callback ) {
  
      if ( i >= secIds.length ) {
        this.emit('end');
        return;
      }
  
      const sectorCallback = (buffer) => {
        if ( bytes - buffer.length < 0 ) {
          buffer = buffer.slice( 0, bytes );
        }
  
        bytes -= buffer.length;
        this.emit('data', buffer);
        callback();
      };
  
      if ( shortStream ) {
        doc._readShortSector( secIds[i], sectorCallback );
      } else {
        doc._readSector( secIds[i], sectorCallback );
      }
    });
  }
  
}

module.exports = Storage;