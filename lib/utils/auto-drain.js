/**
 * Copyright (c) 2015-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

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
