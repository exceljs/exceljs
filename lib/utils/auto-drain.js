'use strict';

var events = require('events');

var utils = require('./utils');

// =============================================================================
// AutoDrain - kind of /dev/null
var AutoDrain = module.exports = function() {
};

utils.inherits(AutoDrain, events.EventEmitter, {
  write: function(chunk) {
    this.emit('data', chunk);
  },
  end: function() {
    this.emit('end');
  }
});
