'use strict';

var verquire = require('../utils/verquire');
var chai = require('chai');

var Excel = verquire('excel');

var expect = chai.expect;

// =============================================================================
// This spec is based around a gold standard Excel workbook 'gold.xlsx'

describe('Gold Book', function() {
  describe('Read', function() {
    var wb;
    before(function() {
      wb = new Excel.Workbook();
      return wb.xlsx.readFile(__dirname + '/data/gold.xlsx');
    });

    it('Values', function() {
      var ws = wb.getWorksheet('Values');

      expect(ws.getCell('B1').value).to.equal('I am Text');
      expect(ws.getCell('B2').value).to.equal(3.14);
      expect(ws.getCell('B3').value).to.equal(5);
      // var b4 = ws.getCell('B4').value;
      // console.log(typeof b4, b4);
      expect(ws.getCell('B4').value).to.equalDate(new Date('2016-05-17T00:00:00.000Z'));
      expect(ws.getCell('B5').value).to.deep.equal({formula: 'B1', result: 'I am Text'});

      expect(ws.getCell('B6').value).to.deep.equal({hyperlink: 'https://www.npmjs.com/package/exceljs', text: 'exceljs'});
    });

    it('Styles', function() {
    });
  });
});
