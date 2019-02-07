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
var ZipReader = function(options) {
  this.count = 0;
  this.jsZip = new JSZip();
  this.stream = new StreamBuf();
  this.stream.on('finish', () => {
    this._process();
  });
  this.getEntryType = options.getEntryType || (() => 'string');
};

utils.inherits(ZipReader, events.EventEmitter, {
  _finished: function() {
    if (!--this.count) {
      PromishLib.Promish.resolve()
        .then(() => {
          this.emit('finished');
        });
    }
  },
  _process: function() {
    var content = this.stream.read();
    this.jsZip.loadAsync(content)
      .then(zip => {
        zip.forEach((path, entry) => {
          if (!entry.dir) {
            this.count++;
            entry.async(this.getEntryType(path))
              .then(data => {
                var entryStream = new StreamBuf();
                entryStream.path = path;
                entryStream.write(data);
                entryStream.autodrain = () => {
                  this._finished();
                };
                entryStream.on('finish', () => {
                  this._finished();
                });

                this.emit('entry', entryStream);
              })
              .catch(error => {
                this.emit('error', error);
              });
          }
        });
      })
      .catch(error => {
        this.emit('error', error);
      });
  },

  // ==========================================================================
  // Stream.Writable interface
  write: function(data, encoding, callback) {
    if (this.error) {
      if (callback) {
        callback(error);
      }
      throw error;
    } else {
      return this.stream.write(data, encoding, callback);
    }
  },
  cork: function() {
    return this.stream.cork();
  },
  uncork: function() {
    return this.stream.uncork();
  },
  end: function() {
    return this.stream.end();
  },
  destroy: function(error) {
    this.emit('finished');
    this.error = error;
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
    if (options.hasOwnProperty('base64') && options.base64) {
      this.zip.file(options.name, data, { base64: true });
    } else {
      this.zip.file(options.name, data);
    }
  },
  finalize: function() {
    var options = {
      type: 'nodebuffer',
      compression: 'DEFLATE'
    };
    return this.zip.generateAsync(options)
      .then(content => {
        this.stream.end(content);
        this.emit('finish');
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
