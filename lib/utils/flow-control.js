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

var events = require('events');

var utils = require('./utils');

// =============================================================================
// FlowControl - Used to slow down streaming to manageable speed
// Implements a subset of Stream.Duplex: pipe() and write()
var FlowControl = module.exports = function(options) {
  this.options = options = options || {};

  // Buffer queue
  this.queue = [];

  // Consumer streams
  this.pipes = [];

  // Down-stream flow-control instances
  this.children = [];

  // Up-stream flow-control instances
  this.parent = options.parent;

  // Ensure we don't flush more than once at a time
  this.flushing = false;

  // determine timeout for flow control delays
  if (options.gc) {
    var gc = options.gc;
    if (gc.getTimeout) {
      this.getTimeout = gc.getTimeout;
    } else {
      // heap size below which we don't bother delaying
      var threshold = gc.threshold !== undefined ? gc.threshold : 150000000;
      // convert from heapsize to ms timeout
      var divisor = gc.divisor !== undefined ? gc.divisor : 500000;
      this.getTimeout = function() {
        var memory = process.memoryUsage();
        var heapSize = memory.heapTotal;
        return (heapSize < threshold) ? 0 : Math.floor(heapSize / divisor);
      };
    }
  } else {
    this.getTimeout = null;
  }
};

utils.inherits(FlowControl, events.EventEmitter, {
  get name() {
    return [
      'FlowControl',
      this.parent ? 'Child' : 'Root',
      this.corked ? 'corked' : 'open'
    ].join(' ');
  },
  get corked() {
    // We remain corked while we have children and at least one has data to consume
    return (this.children.length > 0) &&
      this.children.some(function(child) { return child.queue && child.queue.length; });
  },
  get stem() {
    // the decision to stem the incoming data depends on whether the children are corked
    // and how many buffers we have backed up
    return this.corked || !this.queue || (this.queue.length > 2);
  },

  _write: function(dst, data, encoding) {
    // Write to a single destination and return a promise    
    // console.log(this.name + ' _write: ' + data.length);

    return new Promise(function(resolve, reject) {
      dst.write(data, encoding, function(error) {
        if (error) {
          // console.log(self.name + ' _write error: ' + error.message);            
          reject(error);
        } else {
          // console.log(self.name + ' _write complete');            
          resolve();
        }
      });
    });
  },

  _pipe: function(chunk) {
    var self = this;

    // Write chunk to all pipes. A chunk with no data is the end
    var promises = [];
    this.pipes.forEach(function(pipe) {
      if (chunk.data) {
        if (pipe.sync) {
          pipe.stream.write(chunk.data, chunk.encoding);
        } else {
          promises.push(self._write(pipe.stream, chunk.data, chunk.encoding));
        }
      } else {
        pipe.stream.end();
      }
    });
    if (!promises.length) {
      promises.push(Promise.resolve());
    }
    return Promise.all(promises)
      .then(function() {
        try {
          chunk.callback();
        }
        catch(e) {
        }
      });
  },
  _animate: function() {
    var count = 0;
    var seq = ['|', '/', '-', '\\'];
    //var cr = '\033[0G';
    var cr = '\u001b[0G';
    return setInterval(function() {
      process.stdout.write(seq[count++ % 4] + cr);
    }, 100);
  },
  _delay: function() {
    // in certain situations it may be useful to delay processing (e.g. for GC)
    var timeout = this.getTimeout && this.getTimeout();
    var self = this;
    if (timeout) {
      return new Promise(function(resolve, reject) {
        var anime = self._animate();
        setTimeout(function() {
          clearInterval(anime);
          resolve();
        }, timeout);
      });
    } else {
      return Promise.resolve();
    }
  },

  _flush: function() {
    // If/while not corked and we have buffers to send, send them
    var self = this;
    // console.log(this.name + ' flush:' + (this.queue && this.queue.length));

    if (this.queue && !this.flushing && !this.corked) {
      if (this.queue.length) {
        this.flushing = true;
        self._delay()
          .then(function() {
            return self._pipe(self.queue.shift());
          })
          .then(function() {
            setImmediate(function() {
              self.flushing = false;
              self._flush();
            });
          });
      }

      if (!this.stem) {
        // Signal up-stream that we're ready for more data
        // console.log(self.name + ' emitting drain');
        self.emit('drain');
      }
    }
  },

  write: function(data, encoding, callback) {
    // Called by up-stream pipe
    if (encoding instanceof Function) {
      callback = encoding;
      encoding = 'utf8';
    }
    callback = callback || utils.nop;

    if (!this.queue) {
      throw new Error('Cannot write to stream after end');
    }

    // Always queue chunks and then flush
    // console.log(this.name + ' write: ' + data.length);
    this.queue.push({
      data: data,
      encoding: encoding,
      callback: callback
    });
    this._flush();

    // restrict further incoming data if we have backed up buffers or
    // the children are still busy
    var stemFlow = this.corked || (this.queue.length > 3);
    return !stemFlow;
  },

  end: function(chunk, encoding, callback) {
    // Signal from up-stream 
    var self = this;
    // console.log(this.name + 'End')
    this.queue.push({
      callback: function() {
        self.queue = null;
        self.emit('finish');
      }
    });
    this._flush();
  },

  pipe: function(stream, options) {
    options = options || {};

    // some streams don't call callbacks
    var sync = options.sync || false;

    this.pipes.push({
      stream: stream,
      sync: sync
    });
  },
  unpipe: function(stream) {
    this.pipes = this.pipes.filter(function(pipe) {
      return pipe.stream !== stream;
    });
  },

  createChild: function() {
    // Create a down-stream flow-control
    var self = this;
    var options = Object.assign({parent: this}, this.options);
    var child = new FlowControl(options);
    this.children.push(child);

    child.on('drain', function() {
      // a child is ready for more
      self._flush();
    });

    child.on('finish', function() {
      // One child has finished it's stream. Remove it and continue
      self.children = self.children.filter(function(item) {
        return item !== child;
      });
      self._flush();
    });

    return child;
  }
});
