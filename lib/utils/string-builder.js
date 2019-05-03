'use strict';

// StringBuf - a way to keep string memory operations to a minimum
// while building the strings for the xml files
var StringBuf = module.exports = function() {
  this.reset();
};

StringBuf.prototype = {
  get length() {
    return this._buf.length;
  },
  toString: function() {
    return this._buf.join('');
  },

  reset: function(position) {
    if (position) {
      while (this._buf.length > position) {
        this._buf.pop();
      }
    } else {
      this._buf = [];
    }
  },
  addText: function(text) {
    this._buf.push(text);
  },

  addStringBuf: function(inBuf) {
    this._buf.push(inBuf.toString());
  }
};
