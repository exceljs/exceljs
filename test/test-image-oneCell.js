var fs = require('fs');
var path = require('path');

var HrStopwatch = require('./utils/hr-stopwatch');

var Workbook = require('../excel').Workbook;

var filename = process.argv[2];

var wb = new Workbook();
var ws = wb.addWorksheet('blort');

var imageId = wb.addImage({
  filename: path.join(__dirname, 'data/image2.png'),
  extension: 'png',
});
ws.addImage(imageId, {
  tl: { col: 0.1125, row: 0.4 },
  br: { col: 2.101046875, row: 3.4 },
  editAs: 'oneCell'
});

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
