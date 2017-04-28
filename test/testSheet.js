var fs = require('fs');
var _ = require('underscore');
var Promise = require('bluebird');

var SharedStrings = require('../lib/utils/shared-strings');
var Worksheet = require('../lib/doc/worksheet');

var sheetname = process.argv[2];
var stringname = process.argv[3];
var relname = process.argv[4];

var ws = new Worksheet()

ws.columns = [
  { header: 'Col 1', key:'key', width: 25 },
  { header: 'Col 2', key:'name', width: 25 },
  { header: 'Col 3', key:'age' },
  { header: 'Col 4', key:'addr1', width: 8 },
  { header: 'Col 5', key:'addr2', width: 10 },
];

ws.getCell('A2').value = 'Hello, World!';
ws.getCell('B1').value = 'Hello, World!';
ws.getCell('C1').value = 5;
ws.getCell('D1').value = 3.14;
ws.getCell('C3').value = 'Boo!';

ws.getCell('A4').value = 'merge 3x1';
ws.getCell('B4').value = "Won't see this";
ws.mergeCells('A4:C4');

ws.getCell('B5').value = 'merge 3x3';
ws.mergeCells('B5', 'D7');

ws.getCell('C8').value = 'merge 1x2';
ws.mergeCells(8,3,9,3);

ws.getCell('A10').value = {text:'www.google.com', hyperlink: 'http://www.google.com'};

var promises = [];

console.log('Writing sheet to ' + sheetname);
var sheetstream = fs.createWriteStream(sheetname);
promises.push(ws.write(sheetstream));

console.log('Writing string table to ' + stringname);
var stringstream = fs.createWriteStream(stringname);
promises.push(ws.sharedStrings.write(stringstream));

console.log('Writing relationship table to ' + relname);
var relstream = fs.createWriteStream(relname);
promises.push(ws.relationships.write(relstream));


Promise.all(promises)
  .then(function(){
    sheetstream.close();
    stringstream.close();
    relstream.close();
    console.log('Done.');
  });
