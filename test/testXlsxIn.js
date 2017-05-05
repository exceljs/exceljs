'use strict';

var HrStopwatch = require('./utils/hr-stopwatch');
var Excel = require('../excel');

var filename = process.argv[2];
var wb = new Excel.Workbook();

var stopwatch = new HrStopwatch();
stopwatch.start();

wb.xlsx.readFile(filename)
  .then(function() {
    var micros = stopwatch.microseconds;

    console.log('Loaded', filename);
    console.log('Time taken:', micros / 1000000);

    wb.eachSheet(function(sheet, id) {
      console.log(id, sheet.name);
    });
  })
  .catch(function(error) {
    console.error('something went wrong', error.stack);
  });
