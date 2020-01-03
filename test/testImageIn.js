var fs = require('fs');

var HrStopwatch = require('./utils/hr-stopwatch');

var Workbook = require('../lib/doc/workbook');

var filename = process.argv[2];

var wb = new Workbook();
var stopwatch = new HrStopwatch();
stopwatch.start();
wb.xlsx.readFile(filename)
  .then(function(){
    var micros = stopwatch.microseconds;
    console.log("Done.");
    console.log("Time taken:", micros);

    var ws = wb.getWorksheet("blort");

    var image = ws.background.image;
    console.log('Media', name, image.type, image.buffer.length);
  })
  .catch(function(error) {
     console.log(error.message);
  });
