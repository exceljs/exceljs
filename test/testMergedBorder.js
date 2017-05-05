var fs = require('fs');
var _ = require('underscore');
var Promise = require('bluebird');

var Workbook = require('../lib/doc/workbook');

var filename = process.argv[2];

var wb = new Workbook()
var ws = wb.addWorksheet('blort');


var borders = {
  thin: { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}},
  doubleRed: { color: {argb:'FFFF0000'}, top: {style:'double'}, left: {style:'double'}, bottom: {style:'double'}, right: {style:'double'}}
};

var fills = {
  redDarkVertical: {type: 'pattern', pattern:'darkVertical', fgColor:{argb:'FFFF0000'}},
  redGreenDarkTrellis: {type: 'pattern', pattern:'darkTrellis', fgColor:{argb:'FFFF0000'}, bgColor:{argb:'FF00FF00'}},
  blueWhiteHGrad: {type: 'gradient', gradient: 'angle', degree: 0,
    stops: [{position:0, color:{argb:'FF0000FF'}},{position:1, color:{argb:'FFFFFFFF'}}]},
  rgbPathGrad: {type: 'gradient', gradient: 'path', center:{left:0.5,top:0.5},
    stops: [{position:0, color:{argb:'FFFF0000'}},{position:0.5, color:{argb:'FF00FF00'}},{position:1, color:{argb:'FF0000FF'}}]}
};


ws.getCell('B2').value = 'Hello';
ws.mergeCells('B2:C2');
ws.getCell('B2').border = borders.thin;

ws.getCell('E2').value = 'World';
ws.mergeCells('E2:F3');
ws.getCell('E2').border = borders.thin;
ws.getCell('F2').border = borders.thin;
ws.getCell('E3').border = borders.thin;
ws.getCell('F3').border = borders.thin;

ws.getCell('H2').value = 'Broke';
ws.getCell('H2').border = borders.thin;
ws.mergeCells('H2:I3');
ws.getCell('I3').style = {};

wb.xlsx.writeFile(filename)
  .then(function() {
    console.log('Done.');
  })
  .catch(function(error) {
    console.log(error.message);
  });
