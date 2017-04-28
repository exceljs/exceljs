var fs = require('fs');

var HrStopwatch = require('./utils/hr-stopwatch');

var Workbook = require('../excel').Workbook;

var filenameIn = process.argv[2];
var filenameOut = process.argv[3];

var stopwatch = new HrStopwatch();
var wb = new Workbook();
stopwatch.start();
wb.xlsx.readFile(filenameIn)
  .then(function() {
    return wb.xlsx.writeFile(filenameOut);
  })
  .then(function() {
    var micros = stopwatch.microseconds;
    console.log('Done.');
    console.log('Time taken:', micros)
  })
  .catch(function(error) {
    console.error('Error', error.stack);
  });
