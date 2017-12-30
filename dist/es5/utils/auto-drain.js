/**
 * Copyright (c) 2015-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var events = require('events');

var utils = require('./utils');

// =============================================================================
// AutoDrain - kind of /dev/null
var AutoDrain = module.exports = function () {};

utils.inherits(AutoDrain, events.EventEmitter, {
  write: function write(chunk) {
    this.emit('data', chunk);
  },
  end: function end() {
    this.emit('end');
  }
});
//# sourceMappingURL=auto-drain.js.map
