'use strict';

const events = require('events');

const utils = require('../utils/utils');

const LineBuffer = (module.exports = function(options) {
  events.EventEmitter.call(this);
  this.encoding = options.encoding;

  this.buffer = null;

  // part of cork/uncork
  this.corked = false;
  this.queue = [];
});

utils.inherits(LineBuffer, events.EventEmitter, {
  // Events:
  //  line: here is a line
  //  done: all lines emitted

  write(chunk) {
    // find line or lines in chunk and emit them if not corked
    // or queue them if corked
    const data = this.buffer ? this.buffer + chunk : chunk;
    const lines = data.split(/\r?\n/g);

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
  cork() {
    this.corked = true;
  },
  uncork() {
    this.corked = false;
    this._flush();

    // tell the source I'm ready again
    this.emit('drain');
  },
  setDefaultEncoding() {
    // ?
  },
  end() {
    if (this.buffer) {
      this.emit('line', this.buffer);
      this.buffer = null;
    }
    this.emit('done');
  },

  _flush() {
    if (!this.corked) {
      this.queue.forEach(function(line) {
        this.emit('line', line);
      });
      this.queue = [];
    }
  },
});
