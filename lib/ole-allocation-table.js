
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
  while ( secId > AllocationTable.SecIdFree ) {
    secIds.push( secId );
    var secIdPrior = secId;
    secId = this._table[secId];
    if (secId === secIdPrior) { // this will cause a deadlock and a out of memory error
      break;
    }
  }

  return secIds;
};

module.exports = AllocationTable;
