/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

// The purpose of this module is to wrap the js-zip library into a streaming zip library
// since most of the exceljs code uses streams.
// One day I might find (or build) a properly streaming browser safe zip lib

var events = require('events');
var PromishLib = require('./promish');

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
      PromishLib.Promish.resolve()
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
                self.emit('error', error);
              });
          }
        });
      })
      .catch(function (error) {
        self.emit('error', error);
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
        self.stream.end(content);
        self.emit('finish');
      });
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
    return this.stream.unshift(chunk);
  },
  wrap: function(stream) {
    return this.stream.wrap(stream);
  }
});

// =============================================================================

module.exports = {
  ZipReader: ZipReader,
  ZipWriter: ZipWriter
};