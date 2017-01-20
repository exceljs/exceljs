/**
 * Copyright (c) 2016 Guyon Roche
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

// The purpose of this module is to wrap the js-zip library into a streaming zip library
// since most of the exceljs code uses streams.
// One day I might find (or build) a properly streaming browser safe zip lib

var Stream = require('stream');
var events = require('events');

var JSZip = require('jszip');
var packer = require('zip-stream');
var Bluebird = require('bluebird');

var utils = require('./utils');
var StreamBuf = require('./stream-buf');


// =============================================================================
// The ZipReader class
// Unpacks an incoming zip stream
var ZipReader = function() {
  var self = this;
  this.count = 0;
  this.jsZip = new JSZip();
  this.stream = new StreamBuf();
  this.stream.on('finish', function() {
    self._process();
  });
};

utils.inherits(ZipReader, events.EventEmitter, {
  _finished: function() {
    var self = this;
    if (!--this.count) {
      Bluebird.resolve()
        .then(function() {
          self.emit('finished');
        });
    }
  },
  _process: function() {
    var self = this;
    var content = this.stream.read();
    this.jsZip.loadAsync(content)
      .then(function(zip) {
        zip.forEach(function(path, entry) {
          self.count++;
          entry.async('string')
            .then(function(data) {
              var entryStream = new StreamBuf();
              entryStream.path = path;
              // console.log('data', path, data.toString())
              entryStream.write(data);
              entryStream.autodrain = function() {
                self._finished();
              };
              entryStream.on('finish', function() {
                self._finished();
              });

              self.emit('entry', entryStream);
            })
            .catch(function(error) {
              console.error('caught error', error.stack)
            });
        });
      });
  },

  // ==========================================================================
  // Stream.Writable interface
  write: function(data, encoding, callback) {
    return this.stream.write(data, encoding, callback);
  },
  cork: function() {
    return this.stream.cork();
  },
  uncork: function() {
    return this.stream.uncork();
  },
  end: function() {
    return this.stream.end();
  }
});

// =============================================================================
// The ZipWriter class
// Packs streamed data into an output zip stream
var ZipWriter = function() {
  this.zip = new packer();

  // packer supports only one operation at a time so we need to manage
  // multiple entries
  this.queue = [];
  this.current = null;
  this.finished = false;

  var self = this;
  // this.stream.on('readable', function() { self.emit('readable')} );
  // this.stream.on('data', function(data) { self.emit('data', data)} );
  this.zip.on('finish', function() { self.emit('finish')} );
  this.zip.on('error', function(error) { self.emit('error', error)} );
};

utils.inherits(ZipWriter, events.EventEmitter, {
  _process: function() {
    var self = this;

    if (this.current) {
      return;
    }

    if (this.queue.length) {
      this.current = this.queue.shift();
      console.log('Adding entry', this.current.options.name)
      this.zip.entry(this.current.data, this.current.options, function(error, entry) {
        self.current = null;
        if (error) {
          self.emit('error', error);
        }
        self._process();
      });
    } else if (this.finished) {
      this.zip.finish();
    }
  },
  append: function(data, options) {
    this.queue.push({data: data, options: options});
    this._process();
  },
  finalize: function() {
    this.finished = true;
    this._process();
  },

  // ==========================================================================
  // Stream.Readable interface
  read: function(size) {
    return this.zip.read(size);
  },
  setEncoding: function(encoding) {
    return this.zip.setEncoding(encoding);
  },
  pause: function() {
    return this.zip.pause();
  },
  resume: function() {
    return this.zip.resume();
  },
  isPaused: function() {
    return this.zip.isPaused();
  },
  pipe: function(destination, options) {
    return this.zip.pipe(destination, options);
  },
  unpipe: function(destination) {
    return this.zip.unpipe(destination);
  },
  unshift: function(chunk) {
    return this.zip.unshift(size);
  },
  wrap: function(stream) {
    return this.zip.wrap(size);
  }
});

// =============================================================================

module.exports = {
  ZipReader: ZipReader,
  ZipWriter: ZipWriter
}