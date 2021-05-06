var _ = require('underscore');

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
  }

  if ( storageEntry.storageDirId > -1 ) {
    childIds.push( storageEntry.storageDirId );
    var rootChildEntry = self._entries[storageEntry.storageDirId];
    visit( rootChildEntry );
  }

  return childIds;
};

module.exports = DirectoryTree;
