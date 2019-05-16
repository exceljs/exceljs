'use strict';

const PromiseLib = require('../utils/promise');

function setValue(key, value, overwrite) {
  if (overwrite === undefined) {
    // only avoid overwrite if explicitly disabled
    overwrite = true;
  }
  switch (key.toLowerCase()) {
    case 'promise':
      if (!overwrite && PromiseLib.Promise) return;
      PromiseLib.Promise = value;
      break;
    default:
      break;
  }
}

module.exports = setValue;
