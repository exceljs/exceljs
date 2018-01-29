'use strict';

var expect = require('chai').expect;

var Excel = require('../../../excel');

var simpleWorkbookModel = require('./../data/simpleWorkbook.json');

// =============================================================================
// Helpers

function createSimpleWorkbook() {
  var wb = new Excel.Workbook();
  var ws = wb.addWorksheet('blort');

  // plain number
  ws.getCell('A1').value = 7;
  ws.getCell('A1').name = 'Seven';

  // simple string
  ws.getCell('B1').value = 'Hello, World!';
  ws.getCell('B1').name = 'Hello';

  // floating point
  ws.getCell('C1').value = 3.14;

  // date-time
  ws.getCell('D1').value = new Date();

  // hyperlink
  ws.getCell('E1').value = {text: 'www.google.com', hyperlink: 'http://www.google.com'};

  // number formula
  ws.getCell('A2').value = {formula: 'A1', result: 7};
  ws.getCell('A2').name = 'TheFormula';

  // string formula
  ws.getCell('B2').value = {formula: 'CONCATENATE("Hello", ", ", "World!")', result: 'Hello, World!'};
  ws.getCell('B2').name = 'TheFormula';

  // date formula
  ws.getCell('C2').value = {formula: 'D1', result: new Date()};

  return wb;
}

// =============================================================================
// Tests

describe('Workbook', function() {
  it('stores shared string values properly', function() {
    var wb = new Excel.Workbook();
    var ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';

    ws.getCell('A2').value = 'Hello';
    ws.getCell('B2').value = 'World';
    ws.getCell('C2').value = {formula: 'CONCATENATE(A2, ", ", B2, "!")', result: 'Hello, World!'};

    ws.getCell('A3').value = ['Hello', 'World'].join(', ') + '!';

    // A1 and A3 should reference the same string object
    expect(ws.getCell('A1').value).to.equal(ws.getCell('A3').value);

    // A1 and C2 should not reference the same object
    expect(ws.getCell('A1').value).to.equal(ws.getCell('C2').value.result);
  });

  it('assigns cell types properly', function() {
    var wb = createSimpleWorkbook();
    var ws = wb.getWorksheet('blort');

    expect(ws.getCell('A1').type).to.equal(Excel.ValueType.Number);
    expect(ws.getCell('B1').type).to.equal(Excel.ValueType.String);
    expect(ws.getCell('C1').type).to.equal(Excel.ValueType.Number);
    expect(ws.getCell('D1').type).to.equal(Excel.ValueType.Date);
    expect(ws.getCell('E1').type).to.equal(Excel.ValueType.Hyperlink);

    expect(ws.getCell('A2').type).to.equal(Excel.ValueType.Formula);
    expect(ws.getCell('B2').type).to.equal(Excel.ValueType.Formula);
    expect(ws.getCell('C2').type).to.equal(Excel.ValueType.Formula);
  });

  it('assigns rich text', function() {
    var wb = new Excel.Workbook();
    var ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = {
      richText: [
        {font: {size: 12, color: {theme: 0}, name: 'Calibri', family: 2, scheme: 'minor'}, text: 'This is '},
        {font: {italic: true, size: 12, color: {theme: 0}, name: 'Calibri', scheme: 'minor'}, text: 'a'},
        {font: {size: 12, color: {theme: 1}, name: 'Calibri', family: 2, scheme: 'minor'}, text: ' '},
        {font: {size: 12, color: {argb: 'FFFF6600'}, name: 'Calibri', scheme: 'minor'}, text: 'colorful'},
        {font: {size: 12, color: {theme: 1}, name: 'Calibri', family: 2, scheme: 'minor'}, text: ' text '},
        {font: {size: 12, color: {argb: 'FFCCFFCC'}, name: 'Calibri', scheme: 'minor'}, text: 'with'},
        {font: {size: 12, color: {theme: 1}, name: 'Calibri', family: 2, scheme: 'minor'}, text: ' in-cell '},
        {font: {bold: true, size: 12, color: {theme: 1}, name: 'Calibri', family: 2, scheme: 'minor'}, text: 'format'}
      ]
    };

    expect(ws.getCell('A1').text).to.equal('This is a colorful text with in-cell format');
    expect(ws.getCell('A1').type).to.equal(Excel.ValueType.RichText);
  });

  it.skip('serialises to model', function() {
    var wb = createSimpleWorkbook();
    expect(wb.model).to.deep.equal(simpleWorkbookModel);
  });

  it('returns undefined for non-existant sheet', function() {
    var wb = new Excel.Workbook();
    wb.addWorksheet('first');
    expect(wb.getWorksheet('w00t')).to.equal(undefined);
  });

  it('returns undefined for sheet 0', function() {
    var wb = new Excel.Workbook();
    wb.addWorksheet('first');
    expect(wb.getWorksheet(0)).to.equal(undefined);
  });
  
  it('returns undefined for sheet 0 after accessing wb.worksheets or wb.eachSheet ', function() {
    var wb = new Excel.Workbook();
    var sheet = wb.addWorksheet('first');
    
    wb.eachSheet(function(){});
    var numSheets = wb.worksheets.length;
    
    expect(numSheets).to.equal(1);
    expect(wb.getWorksheet(0)).to.equal(undefined);
    expect(wb.getWorksheet(1) === sheet).to.equal(true);
  });
});
