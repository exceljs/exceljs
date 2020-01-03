// var semver = require('semver');
//
// var useSource = semver.gte(process.version, 'v6.0.0');

var useSource = (process.env.EXCEL_NATIVE === 'yes');

// this function allows the specs to switch between source code and
// transpiled code depending on the environment variable EXCEL_NATIVE
module.exports = function verquire(path) {
  if (path === 'excel') {
    return useSource ?
      require('../../lib/exceljs.nodejs') : // eslint-disable-line
      require('../../dist/es5');  // eslint-disable-line
  }
  return useSource ?
    require('../../lib/' + path) : // eslint-disable-line
    require('../../dist/es5/' + path);  // eslint-disable-line
};
