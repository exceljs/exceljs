/* eslint-disable import/no-extraneous-dependencies,node/no-unpublished-require */
require('regenerator-runtime/runtime');

const ExcelJS = {
  Workbook: require('./doc/workbook'),
};

// Object.assign mono-fill
const Enums = require('./doc/enums');

Object.keys(Enums).forEach(key => {
  ExcelJS[key] = Enums[key];
});

module.exports = ExcelJS;
