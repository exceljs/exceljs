'use strict';

var expect = require('chai').expect;
var bluebird = require('bluebird');
var fs = require('fs');
var fsa = bluebird.promisifyAll(fs);
var _ = require('underscore');
var Excel = require('../../excel');
var testUtils = require('./../testutils');
var utils = require('../../lib/utils/utils');

describe('WorkbookWriter', function() {

  it('creates sheets with correct names', function() {
    var wb = new Excel.stream.xlsx.WorkbookWriter();
    var ws1 = wb.addWorksheet('Hello, World!');
    expect(ws1.name).to.equal('Hello, World!');

    var ws2 = wb.addWorksheet();
    expect(ws2.name).to.match(/sheet\d+/);
  });

  describe('Serialise', function() {
    after(function() {
      // delete the working file, don't care about errors
      return fsa.unlinkAsync('./wbw.test.xlsx').catch(function(){});
    });

    it('xlsx file', function() {
      var options = {
        filename: './wbw.test.xlsx',
        useStyles: true
      };
      var wb = testUtils.createTestBook(true, Excel.stream.xlsx.WorkbookWriter, options);
      //fs.writeFileSync('./testmodel.json', JSON.stringify(wb.model, null, '    '));

      return wb.commit()
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile('./wbw.test.xlsx');
        })
        .then(function(wb2) {
          testUtils.checkTestBook(wb2, 'xlsx', true);
        });
    });

    it('Without styles', function() {
      var options = {
        filename: './wbw.test.xlsx',
        useStyles: false
      };
      var wb = testUtils.createTestBook(true, Excel.stream.xlsx.WorkbookWriter, options);
      //fs.writeFileSync('./testmodel.json', JSON.stringify(wb.model, null, '    '));

      return wb.commit()
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile('./wbw.test.xlsx');
        })
        .then(function(wb2) {
          testUtils.checkTestBook(wb2, 'xlsx', false);
        });
    });

    it('serializes row styles and columns properly', function() {
      var options = {
        filename: './wbw.test.xlsx',
        useStyles: true
      };
      var wb = new Excel.stream.xlsx.WorkbookWriter(options);
      var ws = wb.addWorksheet('blort');

      var colStyle = {
        font: testUtils.styles.fonts.comicSansUdB16,
        alignment: testUtils.styles.namedAlignments.middleCentre
      };
      ws.columns = [
        { header: 'A1', width: 10 },
        { header: 'B1', width: 20, style: colStyle },
        { header: 'C1', width: 30 },
      ];

      ws.getRow(2).font = testUtils.styles.fonts.broadwayRedOutline20;

      ws.getCell('A2').value = 'A2';
      ws.getCell('B2').value = 'B2';
      ws.getCell('C2').value = 'C2';
      ws.getCell('A3').value = 'A3';
      ws.getCell('B3').value = 'B3';
      ws.getCell('C3').value = 'C3';

      return wb.commit()
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile('./wbw.test.xlsx');
        })
        .then(function(wb2) {
          var ws2 = wb2.getWorksheet('blort');
          _.each(['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3'], function(address) {
            expect(ws2.getCell(address).value).to.equal(address);
          });
          expect(ws2.getCell('B1').font).to.deep.equal(testUtils.styles.fonts.comicSansUdB16);
          expect(ws2.getCell('B1').alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);
          expect(ws2.getCell('A2').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
          expect(ws2.getCell('B2').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
          expect(ws2.getCell('C2').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
          expect(ws2.getCell('B3').font).to.deep.equal(testUtils.styles.fonts.comicSansUdB16);
          expect(ws2.getCell('B3').alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);

          expect(ws2.getColumn(2).font).to.deep.equal(testUtils.styles.fonts.comicSansUdB16);
          expect(ws2.getColumn(2).alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);

          expect(ws2.getRow(2).font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
        });
    });

    it('A lot of sheets', function() {
      this.timeout(5000);

      var i;
      var wb = new Excel.stream.xlsx.WorkbookWriter({filename: './wbw.test.xlsx'});
      var numSheets = 90;
      // add numSheets sheets
      for (i = 1; i <= numSheets; i++) {
        var ws = wb.addWorksheet('sheet' + i);
        ws.getCell('A1').value = i;
      }
      return wb.commit()
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile('./wbw.test.xlsx');
        })
        .then(function(wb2) {
          for (i = 1; i <= numSheets; i++) {
            var ws2 = wb2.getWorksheet('sheet' + i);
            expect(ws2).to.be.ok;
            expect(ws2.getCell('A1').value).to.equal(i);
          }
        });
    });

    it('defined names', function() {
      var wb = new Excel.stream.xlsx.WorkbookWriter({filename: './wbw.test.xlsx'});
      var ws = wb.addWorksheet('blort');
      ws.getCell('A1').value = 5;
      ws.getCell('A1').name = 'five';

      ws.getCell('A3').value = 'drei';
      ws.getCell('A3').name = 'threes';
      ws.getCell('B3').value = 'trois';
      ws.getCell('B3').name = 'threes';
      ws.getCell('B3').value = 'san';
      ws.getCell('B3').name = 'threes';

      ws.getCell('E1').value = 'grün';
      ws.getCell('E1').name = 'greens';
      ws.getCell('E2').value = 'vert';
      ws.getCell('E2').name = 'greens';
      ws.getCell('E3').value = 'verde';
      ws.getCell('E3').name = 'greens';

      return wb.commit()
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile('./wbw.test.xlsx');
        })
        .then(function(wb2) {
          var ws2 = wb2.getWorksheet('blort');
          expect(ws2.getCell('A1').name).to.equal('five');

          expect(ws2.getCell('A3').name).to.equal('threes');
          expect(ws2.getCell('B3').name).to.equal('threes');
          expect(ws2.getCell('B3').name).to.equal('threes');

          expect(ws2.getCell('E1').name).to.equal('greens');
          expect(ws2.getCell('E2').name).to.equal('greens');
          expect(ws2.getCell('E3').name).to.equal('greens');
        });
    });

    it('serializes and deserializes dataValidations', function() {
      var wb = new Excel.stream.xlsx.WorkbookWriter({filename: './wbw.test.xlsx'});
      testUtils.addDataValidationSheet(wb);

      return wb.commit()
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile('./wbw.test.xlsx');
        })
        .then(function(wb2) {
          testUtils.checkDataValidationSheet(wb2);
        });
    });


  });
});
