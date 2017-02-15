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
      Promise.resolve()
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
          if (!entry.dir) {
            self.count++;
            entry.async('string')
              .then(function (data) {
                var entryStream = new StreamBuf();
                entryStream.path = path;
                // console.log('data', path, data.toString())
                entryStream.write(data);
                entryStream.autodrain = function () {
                  self._finished();
                };
                entryStream.on('finish', function () {
                  self._finished();
                });

                self.emit('entry', entryStream);
              })
              .catch(function (error) {
                console.error('caught error', error.stack)
              });
          }
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
  this.zip = new JSZip();
  this.stream = new StreamBuf();

  var self = this;
};

utils.inherits(ZipWriter, events.EventEmitter, {

  append: function(data, options) {
    this.zip.file(options.name, data);
  },
  finalize: function() {
    var self = this;

    var options = {
      type: 'nodebuffer',
      compression: 'DEFLATE'
    };
    return this.zip.generateAsync(options)
      .then(function(content) {
        // console.log('zip finalize', typeof content, content)
        self.stream.end(content);
        self.emit('finish');
      });

    // return new Promise(function (resolve) {
    //   self.zip.generateNodeStream(options)
    //     .pipe(self.stream)
    //     .on('finish', function() {
    //       resolve();
    //     });
    // });
  },

  // ==========================================================================
  // Stream.Readable interface
  read: function(size) {
    return this.stream.read(size);
  },
  setEncoding: function(encoding) {
    return this.stream.setEncoding(encoding);
  },
  pause: function() {
    return this.stream.pause();
  },
  resume: function() {
    return this.stream.resume();
  },
  isPaused: function() {
    return this.stream.isPaused();
  },
  pipe: function(destination, options) {
    return this.stream.pipe(destination, options);
  },
  unpipe: function(destination) {
    return this.stream.unpipe(destination);
  },
  unshift: function(chunk) {
    return this.stream.unshift(size);
  },
  wrap: function(stream) {
    return this.stream.wrap(size);
  }
});

// =============================================================================

module.exports = {
  ZipReader: ZipReader,
  ZipWriter: ZipWriter
}