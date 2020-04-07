/* eslint-disable import/no-extraneous-dependencies,node/no-unpublished-require */
require('core-js/modules/es.promise');
require('core-js/modules/es.string.includes');
require('core-js/modules/es.object.assign');
require('core-js/modules/es.object.keys');
require('regenerator-runtime/runtime');

const rewritePattern = require('regexpu-core');
const {generateRegexpuOptions} = require('@babel/helper-create-regexp-features-plugin/lib/util');

const {RegExp} = global;
try {
  RegExp('a', 'u');
} catch (err) {
  global.RegExp = function(pattern, flags) {
    if (flags && flags.includes('u')) {
      return new RegExp(rewritePattern(pattern, flags, generateRegexpuOptions({flags, pattern})));
    }
    return new RegExp(pattern, flags);
  };
  global.RegExp.prototype = RegExp;
}

const ExcelJS = {
  Workbook: require('./doc/workbook'),
};

// Object.assign mono-fill
const Enums = require('./doc/enums');

Object.keys(Enums).forEach(key => {
  ExcelJS[key] = Enums[key];
});

module.exports = ExcelJS;
