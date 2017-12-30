'use strict';

var Stream = require('stream');
var utils = require('./utils');

// =============================================================================
// StreamBase64 - A utility to convert to/from base64 stream
// Note: does not buffer data, must be piped
var StreamBuf = module.exports = function () {
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

  write: function write() /* data, encoding */{
    return true;
  },
  cork: function cork() {},
  uncork: function uncork() {},
  end: function end() /* chunk, encoding, callback */{},

  // readable
  // event readable - some data is now available
  // event data - switch to flowing mode - feeds chunks to handler
  // event end - no more data
  // event close - optional, indicates upstream close
  // event error - duh
  read: function read() /* size */{},
  setEncoding: function setEncoding(encoding) {
    // causes stream.read or stream.on('data) to return strings of encoding instead of Buffer objects
    this.encoding = encoding;
  },
  pause: function pause() {},
  resume: function resume() {},
  isPaused: function isPaused() {},
  pipe: function pipe(destination) {
    // add destination to pipe list & write current buffer
    this.pipes.push(destination);
  },
  unpipe: function unpipe(destination) {
    // remove destination from pipe list
    this.pipes = this.pipes.filter(function (pipe) {
      return pipe !== destination;
    });
  },
  unshift: function unshift() /* chunk */{
    // some numpty has read some data that's not for them and they want to put it back!
    // Might implement this some day
    throw new Error('Not Implemented');
  },
  wrap: function wrap() /* stream */{
    // not implemented
    throw new Error('Not Implemented');
  }
});
//# sourceMappingURL=stream-base64.js.map
