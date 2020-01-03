// this module allows the specs to switch between source code and
// transpiled code depending on the environment variable EXCEL_BUILD

/* eslint-disable import/no-dynamic-require */

const libs = {};
const basePath = (function() {
  switch (process.env.EXCEL_BUILD) {
    case 'es5':
      require('core-js/modules/es.promise');
      require('core-js/modules/es.object.assign');
      require('core-js/modules/es.object.keys');
      require('regenerator-runtime/runtime');
      libs.exceljs = require('../../dist/es5');
      return '../../dist/es5/';
    default:
      libs.exceljs = require('../../lib/exceljs.nodejs');
      return '../../lib/';
  }
})();

module.exports = function verquire(path) {
  if (!libs[path]) {
    libs[path] = require(basePath + path);
  }
  return libs[path];
};
