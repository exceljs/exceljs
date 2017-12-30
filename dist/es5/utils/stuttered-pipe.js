/**
 * Copyright (c) 2015-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var events = require('events');

var utils = require('./utils');

// =============================================================================
// StutteredPipe - Used to slow down streaming so GC can get a look in
var StutteredPipe = module.exports = function (readable, writable, options) {
  var self = this;
  options = options || {};

  this.readable = readable;
  this.writable = writable;
  this.bufSize = options.bufSize || 16384;
  this.autoPause = options.autoPause || false;

  this.paused = false;
  this.eod = false;
  this.scheduled = null;

  readable.on('end', function () {
    self.eod = true;
    writable.end();
  });

  // need to have some way to communicate speed of stream
  // back from the consumer
  readable.on('readable', function () {
    if (!self.paused) {
      self.resume();
    }
  });
  this._schedule();
};

utils.inherits(StutteredPipe, events.EventEmitter, {
  pause: function pause() {
    this.paused = true;
  },
  resume: function resume() {
    if (!this.eod) {
      if (this.scheduled !== null) {
        clearImmediate(this.scheduled);
      }
      this._schedule();
    }
  },
  _schedule: function _schedule() {
    var _this = this;

    this.scheduled = setImmediate(function () {
      _this.scheduled = null;
      if (!_this.eod && !_this.paused) {
        var data = _this.readable.read(_this.bufSize);
        if (data && data.length) {
          _this.writable.write(data);

          if (!_this.paused && !_this.autoPause) {
            _this._schedule();
          }
        } else if (!_this.paused) {
          _this._schedule();
        }
      }
    });
  }
});
//# sourceMappingURL=stuttered-pipe.js.map
