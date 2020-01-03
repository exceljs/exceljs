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
ws.addImage(imageId, {
  tl: { col: 1, row: 1 },
  br: { col: 3.5, row: 5.5 }
});
ws.addImage(imageId, 'B7:E12');

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
