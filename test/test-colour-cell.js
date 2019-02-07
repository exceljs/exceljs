var fs = require('fs');

var HrStopwatch = require('./utils/hr-stopwatch');

var Workbook = require('../lib/doc/workbook');

var filename = process.argv[2];

var wb = new Workbook();
var ws = wb.addWorksheet('blort');


var fills = {
  redDarkVertical: {type: 'pattern', pattern:'darkVertical', fgColor:{argb:'FFFF0000'}},
  redGreenDarkTrellis: {type: 'pattern', pattern:'darkTrellis', fgColor:{argb:'FFFF0000'}, bgColor:{argb:'FF00FF00'}},
  blueWhiteHGrad: {type: 'gradient', gradient: 'angle', degree: 0,
    stops: [{position:0, color:{argb:'FF0000FF'}},{position:1, color:{argb:'FFFFFFFF'}}]},
  rgbPathGrad: {type: 'gradient', gradient: 'path', center:{left:0.5,top:0.5},
    stops: [{position:0, color:{argb:'FFFF0000'}},{position:0.5, color:{argb:'FF00FF00'}},{position:1, color:{argb:'FF0000FF'}}]}
};


ws.addRow([1,2,3,4]);
ws.addRow(['one', 'two', 'three', 'four']);
ws.addRow(['une', 'deux', 'trois', 'quatre']);
ws.addRow(['uno', 'due', 'tre', 'quatro']);

ws.getCell('B2').fill = fills.redDarkVertical;

var stopwatch = new HrStopwatch();
stopwatch.start();
wb.xlsx.writeFile(filename)
  .then(function() {
    var micros = stopwatch.microseconds;
    console.log('Done.');
    console.log('Time taken:', micros)
  })
  .catch(function(error) {
     console.log(error.message);
  });
