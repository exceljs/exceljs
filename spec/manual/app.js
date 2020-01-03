'use strict';

/* eslint-disable no-console */

var fs = require('fs');
var express = require('express');
var ExcelJS = require('../../excel');
var StreamBuf = require('../../lib/utils/stream-buf');
var path = require('path');

console.log('Copying bundle.js to public folder');
fs.createReadStream(__dirname + '/../../dist/exceljs.min.js')
  .pipe(fs.createWriteStream(__dirname + '/public/exceljs.min.js'));
fs.createReadStream(__dirname + '/../../dist/exceljs.js')
  .pipe(fs.createWriteStream(__dirname + '/public/exceljs.js'));

var app = express();

app.use('/', express.static(path.join(__dirname, 'public')));

app.post('/api/upload', function(req, res) {
  var wb = new ExcelJS.Workbook();

  var stream = new StreamBuf();
  stream.on('finish', function() {
    var base64 = stream.read();

    wb.xlsx.load(base64, {base64: true})
      .then(function() {
        var ws = wb.getWorksheet('blort');

        console.log('XLSX uploaded:');
        console.log('A1', ws.getCell('A1').value);
        console.log('A2', ws.getCell('A2').value);

        ws.getCell('A1').value = 'Hey Ho!';
        ws.getCell('A2').value = 14;

        var outStream = new StreamBuf();
        wb.xlsx.write(outStream)
          .then(function() {
            var b = outStream.read();
            var s = b.toString('base64');
            res.write(s);
            res.end();
          });
      });
  });

  req.pipe(stream);
});

app.listen(3003);
console.log('Listening on port 3003');
