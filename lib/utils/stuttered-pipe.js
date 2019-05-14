'use strict';

const events = require('events');

const utils = require('./utils');

// =============================================================================
// StutteredPipe - Used to slow down streaming so GC can get a look in
const StutteredPipe = (module.exports = function(readable, writable, options) {
  const self = this;
  options = options || {};

  this.readable = readable;
  this.writable = writable;
  this.bufSize = options.bufSize || 16384;
  this.autoPause = options.autoPause || false;

  this.paused = false;
  this.eod = false;
  this.scheduled = null;

  readable.on('end', () => {
    self.eod = true;
    writable.end();
  });

  // need to have some way to communicate speed of stream
  // back from the consumer
  readable.on('readable', () => {
    if (!self.paused) {
      self.resume();
    }
  });
  this._schedule();
});

utils.inherits(StutteredPipe, events.EventEmitter, {
  pause() {
    this.paused = true;
  },
  resume() {
    if (!this.eod) {
      if (this.scheduled !== null) {
        clearImmediate(this.scheduled);
      }
      this._schedule();
    }
  },
  _schedule() {
    this.scheduled = setImmediate(() => {
      this.scheduled = null;
      if (!this.eod && !this.paused) {
        const data = this.readable.read(this.bufSize);
        if (data && data.length) {
          this.writable.write(data);

          if (!this.paused && !this.autoPause) {
            this._schedule();
          }
        } else if (!this.paused) {
          this._schedule();
        }
      }
    });
  },
});
