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
var ZipReader = function ZipReader(options) {
  var _this = this;

  this.count = 0;
  this.jsZip = new JSZip();
  this.stream = new StreamBuf();
  this.stream.on('finish', function () {
    _this._process();
  });
  this.getEntryType = options.getEntryType || function () {
    return 'string';
  };
};

utils.inherits(ZipReader, events.EventEmitter, {
  _finished: function _finished() {
    var _this2 = this;

    if (! --this.count) {
      PromishLib.Promish.resolve().then(function () {
        _this2.emit('finished');
      });
    }
  },
  _process: function _process() {
    var _this3 = this;

    var content = this.stream.read();
    this.jsZip.loadAsync(content).then(function (zip) {
      zip.forEach(function (path, entry) {
        if (!entry.dir) {
          _this3.count++;
          entry.async(_this3.getEntryType(path)).then(function (data) {
            var entryStream = new StreamBuf();
            entryStream.path = path;
            entryStream.write(data);
            entryStream.autodrain = function () {
              _this3._finished();
            };
            entryStream.on('finish', function () {
              _this3._finished();
            });

            _this3.emit('entry', entryStream);
          }).catch(function (error) {
            _this3.emit('error', error);
          });
        }
      });
    }).catch(function (error) {
      _this3.emit('error', error);
    });
  },

  // ==========================================================================
  // Stream.Writable interface
  write: function write(data, encoding, callback) {
    if (this.error) {
      if (callback) {
        callback(error);
      }
      throw error;
    } else {
      return this.stream.write(data, encoding, callback);
    }
  },
  cork: function cork() {
    return this.stream.cork();
  },
  uncork: function uncork() {
    return this.stream.uncork();
  },
  end: function end() {
    return this.stream.end();
  },
  destroy: function destroy(error) {
    this.emit('finished');
    this.error = error;
  }
});

// =============================================================================
// The ZipWriter class
// Packs streamed data into an output zip stream
var ZipWriter = function ZipWriter() {
  this.zip = new JSZip();
  this.stream = new StreamBuf();
};

utils.inherits(ZipWriter, events.EventEmitter, {

  append: function append(data, options) {
    if (options.hasOwnProperty('base64') && options.base64) {
      this.zip.file(options.name, data, { base64: true });
    } else {
      this.zip.file(options.name, data);
    }
  },
  finalize: function finalize() {
    var _this4 = this;

    var options = {
      type: 'nodebuffer',
      compression: 'DEFLATE'
    };
    return this.zip.generateAsync(options).then(function (content) {
      _this4.stream.end(content);
      _this4.emit('finish');
    });
  },

  // ==========================================================================
  // Stream.Readable interface
  read: function read(size) {
    return this.stream.read(size);
  },
  setEncoding: function setEncoding(encoding) {
    return this.stream.setEncoding(encoding);
  },
  pause: function pause() {
    return this.stream.pause();
  },
  resume: function resume() {
    return this.stream.resume();
  },
  isPaused: function isPaused() {
    return this.stream.isPaused();
  },
  pipe: function pipe(destination, options) {
    return this.stream.pipe(destination, options);
  },
  unpipe: function unpipe(destination) {
    return this.stream.unpipe(destination);
  },
  unshift: function unshift(chunk) {
    return this.stream.unshift(chunk);
  },
  wrap: function wrap(stream) {
    return this.stream.wrap(stream);
  }
});

// =============================================================================

module.exports = {
  ZipReader: ZipReader,
  ZipWriter: ZipWriter
};
//# sourceMappingURL=zip-stream.js.map
