var fs = require('fs');

var HrStopwatch = require('./utils/hr-stopwatch');

var Workbook = require('../lib/doc/workbook');

var image = process.argv[2];
var filename = process.argv[3];

var wb = new Workbook();
var ws = wb.addWorksheet("blort");

ws.getCell("B2").value = "Hello, World!";

ws.setBackgroundImage({
  filename: image,
  type: 'png'
});

var stopwatch = new HrStopwatch();
stopwatch.start();
wb.xlsx.writeFile(filename)
  .then(function(){
    var micros = stopwatch.microseconds;
    console.log("Done.");
    console.log("Time taken:", micros)
  });
//.catch(function(error) {
//    console.log(error.message);
//})
