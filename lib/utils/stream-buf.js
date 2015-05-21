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
"use strict";

var events = require("events");

var utils = require("./utils");
var StringBuf = require("./string-buf");

var nop = function() {};

var ReadWriteBuf = function(size) {
    this.size = size;
    // the buffer
    this.buffer = new Buffer(size);
    // read index
    this.iRead = 0;
    // write index
    this.iWrite = 0;
}
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
    write: function(data, offset, length) {
        // write as many bytes from data from optional source offset
        // and return number of bytes written
        var size = Math.min(length, this.size - this.iWrite);
        data.copy(this.buffer, this.iWrite, offset, offset + size);
        this.iWrite += size;
        return size;
    }
};

// StreamBuf - a multi-purpose read-write stream
//  As MemBuf - write as much data as you like. Then call toBuffer() to consolidate
//  As StreamHub - pipe to multiple writables
//  As readable stream - feed data into the writable part and have some other code read from it.
var StreamBuf = module.exports = function(options) {
    options = options || {};
    this.bufSize = options.bufSize || 1024 * 1024;
    this.buffers = [];
    
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
}

utils.inherits(StreamBuf, events.EventEmitter, {
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
    
    write: function(chunk, encoding, callback) {
        if (encoding instanceof Function) {
            callback = encoding;
            encoding = "utf8";
        }
        callback = callback || nop;
        
        if (chunk instanceof StringBuf) {
            // special case - write directly to buffer and 
            var inPos = 0;
            var inLen = chunk.length;
            var inBuffer = chunk.buffer;
            while (inPos < inLen) {
                // find writable buffer
                var buffer = this._getWritableBuffer();
                
                // write some data
                inPos += buffer.write(inBuffer, inPos, inLen - inPos);
            }
            return true;
        }
        
        // coerce chunk into a buffer
        var inBuffer = (chunk instanceof Buffer) ? chunk : new Buffer(chunk, encoding);
        if (!inBuffer.length) {
            return true;
        }
        
        // pipe to anyone listening
        var countdown = 1;
        var readable = true;
        if (!this.corked) {
            countdown += this.pipes.length;
            this.pipes.forEach(function(pipe) {
                pipe.write(inBuffer, function() {
                    if (!--countdown == 0) {
                        process.nextTick(callback);
                    }
                });
                readable = false;
            });
        }
        
        if (!this.paused) {
            this.emit("data", inBuffer);
        }
        if (readable) {
            // only add to internal buffers if no pipes active
            var inPos = 0;
            var inLen = inBuffer.length;
            while (inPos < inLen) {
                // find writable buffer
                var buffer = this._getWritableBuffer();
                
                // write some data
                inPos += buffer.write(inBuffer, inPos, inLen - inPos);
            }
            
            this.emit("readable");
        }
        
        if (!--countdown) {
            // this line gets called when there are no pipes
            process.nextTick(callback);
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
            while (this.buffers.length > 1) {
                var buf = this.buffers.shift();
                var data = buf.toBuffer();
                this.pipes.forEach(function(pipe) {
                    pipe.write(data);
                });
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
                self.emit("finish");
                self.emit("end");
                self.emit("close");
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
        if (!this.paused) {
            this._flush();
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
        throw new Error("Not Implemented");
    },
    wrap: function(stream) {
        // not implemented
        throw new Error("Not Implemented");
    }
});
