const _ = require('underscore');

const DIRECTORY_TREE_ENTRY_TYPE_EMPTY =       0;   // eslint-disable-line no-unused-vars
const DIRECTORY_TREE_ENTRY_TYPE_STORAGE =     1;
const DIRECTORY_TREE_ENTRY_TYPE_STREAM =      2;
const DIRECTORY_TREE_ENTRY_TYPE_ROOT =        5;

const DIRECTORY_TREE_NODE_COLOR_RED =         0;   // eslint-disable-line no-unused-vars
const DIRECTORY_TREE_NODE_COLOR_BLACK =       1;   // eslint-disable-line no-unused-vars

const DIRECTORY_TREE_LEAF =                  -1;

class DirectoryTree {

  constructor(doc) {
    this._doc = doc;
  }

  load( secIds, callback ) {
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
        return entry.type === DIRECTORY_TREE_ENTRY_TYPE_ROOT;
      });
  
      self._buildHierarchy( self.root );
  
      callback();
    });
  }

  _buildHierarchy( storageEntry ) {
    var self = this;
    var childIds = this._getChildIds( storageEntry );
  
    storageEntry.storages = {};
    storageEntry.streams  = {};
  
    _.each( childIds, function( childId ) {
      var childEntry = self._entries[childId];
      var name = childEntry.name;
      if ( childEntry.type === DIRECTORY_TREE_ENTRY_TYPE_STORAGE ) {
        storageEntry.storages[name] = childEntry;
      }
      if ( childEntry.type === DIRECTORY_TREE_ENTRY_TYPE_STREAM ) {
        storageEntry.streams[name] = childEntry;
      }
    });
  
    _.each( storageEntry.storages, function( childStorageEntry ) {
      self._buildHierarchy( childStorageEntry );
    });
  }
  
  _getChildIds( storageEntry ) {
    var self = this;
    var childIds = [];
  
    function visit( visitEntry ) {
      if ( visitEntry.left !== DIRECTORY_TREE_LEAF ) {
        childIds.push( visitEntry.left );
        visit( self._entries[visitEntry.left] );
      }
      if ( visitEntry.right !== DIRECTORY_TREE_LEAF ) {
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
  }

}

module.exports = DirectoryTree;
