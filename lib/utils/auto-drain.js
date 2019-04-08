'use strict';

const events = require('events');

const utils = require('./utils');

// =============================================================================
// AutoDrain - kind of /dev/null
const AutoDrain = (module.exports = function() {});

utils.inherits(AutoDrain, events.EventEmitter, {
  write(chunk) {
    this.emit('data', chunk);
  },
  end() {
    this.emit('end');
  },
});
