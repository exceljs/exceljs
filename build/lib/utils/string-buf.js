/**
 * Copyright (c) 2015-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

// StringBuf - a way to keep string memory operations to a minimum
// while building the strings for the xml files

var StringBuf = module.exports = function (options) {
  this._buf = new Buffer(options && options.size || 16384);
  this._encoding = options && options.encoding || 'utf8';

  // where in the buffer we are at
  this._inPos = 0;

  // for use by toBuffer()
  this._buffer = undefined;
};

StringBuf.prototype = {

  get length() {
    return this._inPos;
  },
  get capacity() {
    return this._buf.length;
  },
  get buffer() {
    return this._buf;
  },

  toBuffer: function toBuffer() {
    // return the current data as a single enclosing buffer
    if (!this._buffer) {
      this._buffer = new Buffer(this.length);
      this._buf.copy(this._buffer, 0, 0, this.length);
    }
    return this._buffer;
  },

  reset: function reset(position) {
    position = position || 0;
    this._buffer = undefined;
    this._inPos = position;
  },

  _grow: function _grow(min) {
    var size = this._buf.length * 2;
    while (size < min) {
      size *= 2;
    }
    var buf = new Buffer(size);
    this._buf.copy(buf, 0);
    this._buf = buf;
  },

  addText: function addText(text) {
    this._buffer = undefined;

    var inPos = this._inPos + this._buf.write(text, this._inPos, this._encoding);

    // if we've hit (or nearing capacity), grow the buf
    while (inPos >= this._buf.length - 4) {
      this._grow(this._inPos + text.length);

      // keep trying to write until we've completely written the text
      inPos = this._inPos + this._buf.write(text, this._inPos, this._encoding);
    }

    this._inPos = inPos;
  },

  addStringBuf: function addStringBuf(inBuf) {
    if (inBuf.length) {
      this._buffer = undefined;

      if (this.length + inBuf.length > this.capacity) {
        this._grow(this.length + inBuf.length);
      }
      // eslint-disable-next-line no-underscore-dangle
      inBuf._buf.copy(this._buf, this._inPos, 0, inBuf.length);
      this._inPos += inBuf.length;
    }
  }
};
//# sourceMappingURL=string-buf.js.map
