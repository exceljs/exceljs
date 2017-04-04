var fs = require('fs');
var path = require('path');

var HrStopwatch = require('./utils/hr-stopwatch');

var Workbook = require('../excel').Workbook;

var image = process.argv[2];
var filename = process.argv[3];

var wb = new Workbook();
var ws = wb.addWorksheet('blort');

ws.getCell('B2').value = 'Hello, World!';

var imageId = wb.addImage({
  filename: image,
  extension: path.extname(image).substr(1),
})
ws.addImage(imageId, 'C2:D3');
ws.addImage(imageId, 'B5:E10');

var stopwatch = new HrStopwatch();
stopwatch.start();
wb.xlsx.writeFile(filename)
  .then(function() {
    var micros = stopwatch.microseconds;
    console.log('Done.');
    console.log('Time taken:', micros)
  })
  .catch(function(error) {
     console.error(error.stack);
  });
