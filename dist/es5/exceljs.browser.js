'use strict';

// for the benefit of browserify, include browser friendly promise
var setConfigValue = require('./config/set-value');
setConfigValue('promise', require('promish/dist/promish-node'), false);

var ExcelJS = {
  Workbook: require('./doc/workbook')
};

// Object.assign mono-fill
var Enums = require('./doc/enums');

Object.keys(Enums).forEach(function (key) {
  ExcelJS[key] = Enums[key];
});

module.exports = ExcelJS;
//# sourceMappingURL=exceljs.browser.js.map
