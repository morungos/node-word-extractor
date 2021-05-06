
const ALLOCATION_TABLE_SEC_ID_FREE               = -1;
const ALLOCATION_TABLE_SEC_ID_END_OF_CHAIN       = -2;   // eslint-disable-line no-unused-vars
const ALLOCATION_TABLE_SEC_ID_SAT                = -3;   // eslint-disable-line no-unused-vars
const ALLOCATION_TABLE_SEC_ID_MSAT               = -4;   // eslint-disable-line no-unused-vars

class AllocationTable {

  constructor(doc) {
    this._doc = doc;
  }

  load(secIds, callback) {
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
  }
  
  getSecIdChain(startSecId) {
    var secId = startSecId;
    var secIds = [];
    while ( secId > ALLOCATION_TABLE_SEC_ID_FREE ) {
      secIds.push( secId );
      var secIdPrior = secId;
      secId = this._table[secId];
      if (secId === secIdPrior) { // this will cause a deadlock and a out of memory error
        break;
      }
    }
  
    return secIds;
  }
  
}

module.exports = AllocationTable;
