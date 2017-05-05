var semver = require('semver');

var useSource = semver.gte(process.version, 'v6.0.0');

// this function requires in the appropriate branch of exceljs based on the node verson
// it allows integration and end-to-end tests to run natively (without build) for node
// >= 6.0.0 and against the compiled dist folder otherwise.
module.exports = function verquire(path) {
  if (path === 'excel') {
    return useSource ?
      require('../../lib/exceljs.nodejs') : // eslint-disable-line
      require('../../dist/es3');  // eslint-disable-line
  }
  return useSource ?
    require('../../lib/' + path) : // eslint-disable-line
    require('../../dist/es3/' + path);  // eslint-disable-line
}