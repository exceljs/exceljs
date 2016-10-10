/**
 * Copyright (c) 2015 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 */
'use strict';

var utils = require('./utils');

// StringBuf - a way to keep string memory operations to a minimum
// while building the strings for the xml files
var StringBuf = module.exports = function(options) {
  this._buf = new Buffer((options && options.size) || 16384);
  this._encoding = (options && options.encoding) || 'utf8';

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

  toBuffer: function() {
    // return the current data as a single enclosing buffer
    if (!this._buffer) {
      this._buffer = new Buffer(this.length);
      this._buf.copy(this._buffer, 0, 0, this.length);
    }
    return this._buffer;
  },

  reset: function(position) {
    position = position || 0;
    this._buffer = undefined;
    this._inPos = position;
  },

  _grow: function(min) {
    for (var size = this._buf.length * 2; size < min; size *= 2);
    //console.log("Grow: min="+min + ", size="+size);
    var buf = new Buffer(size);
    this._buf.copy(buf, 0);
    this._buf = buf;
  },

  addText: function(text) {
    this._buffer = undefined;

    //console.log("addText: writing " + text.length + ' at ' + this._inPos + ", size= "+ this._buf.length);
    var inPos = this._inPos + this._buf.write(text, this._inPos, this._encoding);

    // if we've hit (or nearing capacity), grow the buf
    while (inPos >= this._buf.length - 4) {
      this._grow(this._inPos + text.length);

      // keep trying to write until we've completely written the text
      inPos = this._inPos + this._buf.write(text, this._inPos, this._encoding);
    }

    this._inPos = inPos;
  },

  addStringBuf: function(inBuf) {
    if (inBuf.length) {
      this._buffer = undefined;

      if (this.length + inBuf.length > this.capacity) {
        this._grow(this.length + inBuf.length);
      }
      inBuf._buf.copy(this._buf, this._inPos, 0, inBuf.length);
      this._inPos += inBuf.length;
    }
  }
};
