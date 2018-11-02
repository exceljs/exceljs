/**
 * Copyright (c) 2015-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var events = require('events');

var utils = require('../utils/utils');

var LineBuffer = module.exports = function (options) {
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

  write: function write(chunk) {
    // find line or lines in chunk and emit them if not corked
    // or queue them if corked
    var data = this.buffer ? this.buffer + chunk : chunk;
    var lines = data.split(/\r?\n/g);

    // save the last line
    this.buffer = lines.pop();

    lines.forEach(function (line) {
      if (this.corked) {
        this.queue.push(line);
      } else {
        this.emit('line', line);
      }
    });

    return !this.corked;
  },
  cork: function cork() {
    this.corked = true;
  },
  uncork: function uncork() {
    this.corked = false;
    this._flush();

    // tell the source I'm ready again
    this.emit('drain');
  },
  setDefaultEncoding: function setDefaultEncoding() {
    // ?
  },
  end: function end() {
    if (this.buffer) {
      this.emit('line', this.buffer);
      this.buffer = null;
    }
    this.emit('done');
  },

  _flush: function _flush() {
    if (!this.corked) {
      this.queue.forEach(function (line) {
        this.emit('line', line);
      });
      this.queue = [];
    }
  }
});
//# sourceMappingURL=line-buffer.js.map
