/**
 * Copyright (c) 2015-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

// StringBuf - a way to keep string memory operations to a minimum
// while building the strings for the xml files
const StringBuf = (module.exports = function() {
  this.reset();
});

StringBuf.prototype = {
  get length() {
    return this._buf.length;
  },
  toString() {
    return this._buf.join('');
  },

  reset(position) {
    if (position) {
      while (this._buf.length > position) {
        this._buf.pop();
      }
    } else {
      this._buf = [];
    }
  },
  addText(text) {
    this._buf.push(text);
  },

  addStringBuf(inBuf) {
    this._buf.push(inBuf.toString());
  },
};
