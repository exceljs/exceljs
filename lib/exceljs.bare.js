// this bundle is built without polyfill leaving apps the freedom to add their own
require('./utils/browser-patch-buffer');

const ExcelJS = {
  Workbook: require('./doc/workbook'),
};

// Object.assign mono-fill
const Enums = require('./doc/enums');

Object.keys(Enums).forEach(key => {
  ExcelJS[key] = Enums[key];
});

module.exports = ExcelJS;
