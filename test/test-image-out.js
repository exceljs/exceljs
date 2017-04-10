var fs = require('fs');
var path = require('path');

var HrStopwatch = require('./utils/hr-stopwatch');

var Workbook = require('../excel').Workbook;

var filename = process.argv[2];

var wb = new Workbook();
var ws = wb.addWorksheet('blort');

ws.getCell('B2').value = 'Hello, World!';

var imageId = wb.addImage({
  filename: path.join(__dirname, 'data/image2.png'),
  extension: 'png',
});
var backgroundId = wb.addImage({
  buffer: fs.readFileSync(path.join(__dirname, 'data/bubbles.jpg')),
  extension: 'jpeg',
});
ws.addImage(imageId, 'C2:D3');
ws.addImage(imageId, 'B5:E10');

ws.addBackgroundImage(backgroundId);

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
