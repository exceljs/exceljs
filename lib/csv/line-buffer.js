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

var events = require('events');

var utils = require('../utils/utils');


var LineBuffer = module.exports = function(options) {
  events.EventEmitter.call(this);
  this.encoding = options.encoding;

  this.buffer = null;

  // part of cork/uncork
  this.corked = false;
  this.queue = [];
};

utils.inherits(LineBuffer, events.EventEmitter, {
  // Events:
  //  line: here is a line
  //  done: all lines emitted

  write: function(chunk, encoding, callback) {

    // find line or lines in chunk and emit them if not corked
    // or queue them if corked
    var data = this.buffer ? this.buffer + chunk : chunk;
    var lines = data.split(/\r?\n/g);

    // save the last line
    this.buffer = lines.pop();

    lines.forEach(function(line) {
      if (this.corked) {
        this.queue.push(line);
      } else {
        this.emit('line', line);
      }
    });

    return !this.corked;
  },
  cork: function() {
    this.corked = true;
  },
  uncork: function() {
    this.corked = false;
    this._flush();

    // tell the source I'm ready again
    this.emit('drain');
  },
  setDefaultEncoding: function(encoding) {
    // ?
  },
  end: function(chunk, encoding, callback) {
    if (this.buffer) {
      this.emit('line', this.buffer);
      this.buffer = null;
    }
    this.emit('done');
  },

  _flush: function() {
    if (!this.corked) {
      this.queue.forEach(function(line) {
        this.emit('line', line);
      });
      this.queue = [];
    }
  }
});
