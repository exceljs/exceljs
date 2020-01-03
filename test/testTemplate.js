'use strict';

var fs = require('fs');
var _ = require('underscore');
var HrStopwatch = require('./utils/hr-stopwatch');

var Excel = require('../excel');

var inputFile = process.argv[2];
var outputFile = process.argv[3];


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


var workbook = new Excel.Workbook();
workbook.xlsx.readFile('./out/template.xlsx')
  .then(function(stream) {
    var options = {
      useSharedStrings: true,
      useStyles: true
    };

    return stream.xlsx.writeFile('./out/template-out.xlsx', options)
      .then(function(){
        console.log('Done.');
      })
  })
  .catch(function(error) {
    console.error(error.message);
    console.error(error.stack);
  });
