// for the benefit of browserify, include browser friendly promise
const setConfigValue = require('./config/set-value');
setConfigValue('promise', require('promish/dist/promish-node'), false);

const ExcelJS = {
  Workbook: require('./doc/workbook'),
};

// Object.assign mono-fill
const Enums = require('./doc/enums');

Object.keys(Enums).forEach(key => {
  ExcelJS[key] = Enums[key];
});

module.exports = ExcelJS;
