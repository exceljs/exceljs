var fs = require('fs');
var _ = require('underscore');
var Promise = require('bluebird');

var Workbook = require('../lib/doc/workbook');

var filename = process.argv[2];

var workbook = new Workbook();
workbook.xlsx.readFile(filename)
  .then(function() {
    workbook.eachSheet(function(worksheet) {
      console.log('Sheet ' + worksheet.id + ' - ' + worksheet.name + ', Dims=' + JSON.stringify(worksheet.dimensions));
    });
  })
  .catch(function(error) {
    console.log(error.message);
  });
