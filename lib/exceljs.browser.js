// for the benefit of browserify, include browser friendly promise
var setConfigValue = require('./config/set-value');
setConfigValue('promise', require('promish/dist/promish-node'), false);

var ExcelJS = {
  Workbook: require('./doc/workbook')
};

// Object.assign mono-fill
var Enums = require('./doc/enums');

for (var i in Enums) {
  if (Enums.hasOwnProperty(i)) {
    ExcelJS[i] = Enums[i];
  }
}

module.exports = ExcelJS;