'use strict';

var Stream = require('stream');
var PromishLib = require('./promish');

var utils = require('./utils');
var StringBuf = require('./string-buf');


// =============================================================================
// StreamBase64 - A utility to convert to/from base64 stream
// Note: does not buffer data, must be piped
var StreamBuf = module.exports = function(options) {
  options = options || {};

  // consuming pipe streams go here
  this.pipes = [];
};

utils.inherits(StreamBuf, Stream.Duplex, {
  // writable
  // event drain - if write returns false (which it won't), indicates when safe to write again.
  // finish - end() has been called
  // pipe(src) - pipe() has been called on readable
  // unpipe(src) - unpipe() has been called on readable
  // error - duh

  write: function(data, encoding, callback) {
    if (encoding instanceof Function) {
      callback = encoding;
      encoding = 'utf8';
    }
    callback = callback || utils.nop;

    return true;
  },
  cork: function() {
  },
  uncork: function() {
  },
  end: function(chunk, encoding, callback) {
  },

  // readable
  // event readable - some data is now available
  // event data - switch to flowing mode - feeds chunks to handler
  // event end - no more data
  // event close - optional, indicates upstream close
  // event error - duh
  read: function(size) {
  },
  setEncoding: function(encoding) {
    // causes stream.read or stream.on('data) to return strings of encoding instead of Buffer objects
    this.encoding = encoding;
  },
  pause: function() {
  },
  resume: function() {
  },
  isPaused: function() {
  },
  pipe: function(destination, options) {
    // add destination to pipe list & write current buffer
    this.pipes.push(destination);
  },
  unpipe: function(destination) {
    // remove destination from pipe list
    this.pipes = this.pipes.filter(function(pipe) {
      return pipe !== destination;
    });
  },
  unshift: function(chunk) {
    // some numpty has read some data that's not for them and they want to put it back!
    // Might implement this some day
    throw new Error('Not Implemented');
  },
  wrap: function(stream) {
    // not implemented
    throw new Error('Not Implemented');
  }
});
