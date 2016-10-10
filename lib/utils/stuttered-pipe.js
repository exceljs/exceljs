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

var utils = require('./utils');

// =============================================================================
// StutteredPipe - Used to slow down streaming so GC can get a look in
var StutteredPipe = module.exports = function(readable, writable, options) {
  var self = this;
  options = options || {};

  this.readable = readable;
  this.writable = writable;
  this.bufSize = options.bufSize || 16384;
  this.autoPause = options.autoPause || false;

  this.paused = false;
  this.eod = false;
  this.scheduled = null;

  readable.on('end', function() {
    //console.log("Input Stream ended")
    self.eod = true;
    writable.end();
  });

  // need to have some way to communicate speed of stream
  // back from the consumer
  readable.on('readable', function() {
    if (!self.paused) {
      self.resume();
    }
  });
  this._schedule();
};

utils.inherits(StutteredPipe, events.EventEmitter, {
  pause: function() {
    this.paused = true;
  },
  resume: function() {
    if (!this.eod) {
      if (this.scheduled !== null) {
        clearImmediate(this.scheduled);
      }
      this._schedule();
    }
  },
  _schedule: function() {
    var self = this;
    this.scheduled = setImmediate(function() {
      self.scheduled = null;
      if (!self.eod && !self.paused) {
        var data = self.readable.read(self.bufSize);
        if (data && data.length) {
          self.writable.write(data);

          if (!self.paused && !self.autoPause) {
            self._schedule();
          }
        } else {
          if (!self.paused) {
            self._schedule();
          }
        }
      }
    });
  }
});
