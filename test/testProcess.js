'use strict';

var fs = require('fs');
var _ = require('underscore');
var HrStopwatch = require('./utils/hr-stopwatch');

var Excel = require('../excel');

var inputFile = process.argv[2];
var outputFile = process.argv[3];

var wb = new Excel.Workbook();

var passed = true;
var assert = function(value, failMessage, passMessage) {
  if (!value) {
    if (failMessage) {
      console.error(failMessage);
    }
    passed = false;
  } else if (passMessage) {
    console.log(passMessage);
  }
};

// assuming file created by testBookOut
wb.xlsx.readFile(inputFile)
  .then(function() {
    console.log('Loaded', inputFile);

    wb.eachSheet(function(sheet) {
      console.log(sheet.name);
    });

    var ws = wb.getWorksheet('Sheet1');

    assert(ws, 'Expected to find a worksheet called sheet1');

    ws.getCell('B1').value = new Date();
    ws.getCell('B1').numFmt = 'hh:mm:ss';

    ws.addRow([1, 'hello']);
    return wb.xlsx.writeFile(outputFile);
  })
  .then(function() {
    assert(passed, 'Something went wrong', 'All tests passed!');
  })
  .catch(function(error) {
    console.error(error.message);
    console.error(error.stack);
  });
