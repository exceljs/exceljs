'use strict';

var fs = require('fs');
var express = require('express');
var ExcelJS = require('../../excel');
var path = require('path');

fs.createReadStream(__dirname + '/../../bundle.js')
  .pipe(fs.createWriteStream(__dirname + '/public/bundle.js'));

var app = express();


app.use('/', express.static(path.join(__dirname, 'public')));

app.post('/api/upload', function(req, res) {
  var wb = new ExcelJS.Workbook();
  req.pipe(wb.xlsx.createInputStream());

  var ws = wb.getWorksheet('foo');

  console.log('A1', ws.getCell('A1').value);
  res.status(200).send('OK');
});

app.listen(3003);
console.log('Listening on port 3003');