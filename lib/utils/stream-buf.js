/**
 * Copyright (c) 2015 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the 'Software'), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 */
'use strict';

var Stream = require('stream');

var utils = require('./utils');
var StringBuf = require('./string-buf');


// =============================================================================
// data chunks - encapsulating incoming data
var StringChunk = function(data, encoding) {
  this._data = data;
  this._encoding = encoding;
};
StringChunk.prototype = {
  get length() {
    return this.toBuffer().length;
  },
  // copy to target buffer
  copy: function(target, targetOffset, offset, length) {
    return this.toBuffer().copy(target, targetOffset, offset, length);
  },
  toBuffer: function() {
    if (!this._buffer) {
      this._buffer = new Buffer(this._data, this._encoding);
    }
    return this._buffer;
  }
};
var StringBufChunk = function(data) {
  this._data = data;
};
StringBufChunk.prototype = {
  get length() {
    return this._data.length;
  },
  // copy to target buffer
  copy: function(target, targetOffset, offset, length) {
    return this._data._buf.copy(target, targetOffset, offset, length);
  },
  toBuffer: function() {
    return this._data.toBuffer();
  }
};
var BufferChunk = function(data) {
  this._data = data;
};
BufferChunk.prototype = {
  get length() {
    return this._data.length;
  },
  // copy to target buffer
  copy: function(target, targetOffset, offset, length) {
    this._data.copy(target, targetOffset, offset, length);
  },
  toBuffer: function() {
    return this._data;
  }
};

// =============================================================================
// ReadWriteBuf - a single buffer supporting simple read-write
var ReadWriteBuf = function(size) {
  this.size = size;
  // the buffer
  this.buffer = new Buffer(size);
  // read index
  this.iRead = 0;
  // write index
  this.iWrite = 0;
};
ReadWriteBuf.prototype = {
  toBuffer: function() {
    if ((this.iRead === 0) && (this.iWrite === this.size)) {
      return this.buffer;
    } else {
      var buf = new Buffer(this.iWrite - this.iRead);
      this.buffer.copy(buf, 0, this.iRead, this.iWrite);
      return buf;
    }
  },
  get length() {
    return this.iWrite - this.iRead;
  },
  get eod() {
    return this.iRead === this.iWrite;
  },
  get full() {
    return this.iWrite === this.size;
  },
  read: function(size) {
    // read size bytes from buffer and return buffer
    if (size === 0) {
      // special case - return null if no data requested
      return null;
    } else if ((size === undefined) || (size >= this.length)) {
      // if no size specified or size is at least what we have then return all of the bytes
      var buf = this.toBuffer();
      this.iRead = this.iWrite;
      return buf;
    } else {
      // otherwise return a chunk
      var buf = new Buffer(size);
      this.buffer.copy(buf, 0, this.iRead, size);
      this.iRead += size;
      return buf;
    }
  },
  write: function(chunk, offset, length) {
    // write as many bytes from data from optional source offset
    // and return number of bytes written
    var size = Math.min(length, this.size - this.iWrite);
    chunk.copy(this.buffer, this.iWrite, offset, offset + size);
    this.iWrite += size;
    return size;
  }
};

// =============================================================================
// StreamBuf - a multi-purpose read-write stream
//  As MemBuf - write as much data as you like. Then call toBuffer() to consolidate
//  As StreamHub - pipe to multiple writables
//  As readable stream - feed data into the writable part and have some other code read from it.
var StreamBuf = module.exports = function(options) {
  options = options || {};
  this.bufSize = options.bufSize || 1024 * 1024;
  this.buffers = [];

  // batch mode fills a buffer completely before passing the data on
  // to pipes or 'readable' event listeners
  this.batch = options.batch || false;

  this.corked = false;
  // where in the current writable buffer we're up to
  this.inPos = 0;

  // where in the current readable buffer we've read up to
  this.outPos = 0;

  // consuming pipe streams go here
  this.pipes = [];

  // controls emit('data') 
  this.paused = false;

  this.encoding = null;
};

utils.inherits(StreamBuf, Stream.Duplex, {
  // writable
  // event drain - if write returns false (which it won't), indicates when safe to write again.
  // finish - end() has been called
  // pipe(src) - pipe() has been called on readable
  // unpipe(src) - unpipe() has been called on readable
  // error - duh

  _getWritableBuffer: function() {
    if (this.buffers.length) {
      var last = this.buffers[this.buffers.length - 1];
      if (!last.full) {
        return last;
      }
    }
    var buf = new ReadWriteBuf(this.bufSize);
    this.buffers.push(buf);
    return buf;
  },

  _pipe: function(chunk) {
    var write = function(pipe) {
      return new Promise(function(resolve, reject) {
        pipe.write(chunk.toBuffer(), function() {
          resolve();
        });
      });
    };
    var promises = this.pipes.map(write);
    return promises.length ?
      Promise.all(promises).then(utils.nop) :
      Promise.resolve();
  },
  _writeToBuffers: function(chunk) {
    var inPos = 0;
    var inLen = chunk.length;
    while (inPos < inLen) {
      // find writable buffer
      var buffer = this._getWritableBuffer();

      // write some data
      inPos += buffer.write(chunk, inPos, inLen - inPos);
    }
  },
  write: function(data, encoding, callback) {
    if (encoding instanceof Function) {
      callback = encoding;
      encoding = 'utf8';
    }
    callback = callback || utils.nop;

    // encapsulate data into a chunk
    var chunk;
    if (data instanceof StringBuf) {
      chunk = new StringBufChunk(data);
    } else if (data instanceof Buffer) {
      chunk = new BufferChunk(data);
    } else {
      // assume string
      chunk = new StringChunk(data, encoding);
    }

    // now, do something with the chunk
    if (this.pipes.length) {
      if (this.batch) {
        this._writeToBuffers(chunk);
        while (!this.corked && (this.buffers.length > 1)) {
          this._pipe(this.buffers.shift());
        }
      } else {
        if (!this.corked) {
          this._pipe(chunk).then(callback);
        } else {
          this._writeToBuffers(chunk);
          process.nextTick(callback);
        }
      }
    } else {
      if (!this.paused) {
        this.emit('data', chunk.toBuffer());
      }

      this._writeToBuffers(chunk);
      this.emit('readable');
    }

    return true;
  },
  cork: function() {
    this.corked = true;
  },
  _flush: function(destination) {
    // if we have comsumers...
    if (this.pipes.length) {
      // and there's stuff not written
      while (this.buffers.length) {
        this._pipe(this.buffers.shift());
      }
    }
  },
  uncork: function() {
    this.corked = false;
    this._flush();
  },
  end: function(chunk, encoding, callback) {
    var self = this;
    var writeComplete = function(error) {
      if (error) {
        callback(error);
      } else {
        self._flush();
        self.pipes.forEach(function(pipe) {
          pipe.end();
        });
        self.emit('finish');
      }
    };
    if (chunk) {
      this.write(chunk, encoding, writeComplete);
    } else {
      writeComplete();
    }
  },

  // readable
  // event readable - some data is now available
  // event data - switch to flowing mode - feeds chunks to handler
  // event end - no more data
  // event close - optional, indicates upstream close
  // event error - duh
  read: function(size) {
    // read min(buffer, size || infinity)
    if (size) {
      var buffers = [];
      while (size && this.buffers.length && !this.buffers[0].eod) {
        var first = this.buffers[0];
        var buffer = first.read(size);
        size -= buffer.length;
        buffers.push(buffer);
        if (first.eod && first.full) {
          this.buffers.shift();
        }
      }
      return Buffer.concat(buffers);
    } else {
      var buffers = this.buffers.map(function(buffer) {
        return buffer.toBuffer();
      }).filter(function(data) {
        return data;
      });
      this.buffers = [];
      return Buffer.concat(buffers);
    }
  },
  setEncoding: function(encoding) {
    // causes stream.read or stream.on('data) to return strings of encoding instead of Buffer objects
    this.encoding = encoding;
  },
  pause: function() {
    this.paused = true;
  },
  resume: function() {
    this.paused = false;
  },
  isPaused: function() {
    return this.paused ? true : false;
  },
  pipe: function(destination, options) {
    // add destination to pipe list & write current buffer
    this.pipes.push(destination);
    if (!this.paused && this.buffers.length) {
      this.end();
    }
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
