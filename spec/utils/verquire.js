var semver = require('semver');

var useSource = semver.gte(process.version, 'v6.0.0');

// this function requires in the appropriate branch of exceljs based on the node verson
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