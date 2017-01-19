'use strict';

var chai = require('chai');
var expect = chai.expect;
chai.use(require('chai-datetime'));

var Excel = require('../../excel');

// this file to contain integration tests created from github issues
var TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';


describe('github issues', function() {
  it('issue 99 - Too few data or empty worksheet generate malformed excel file', function() {
      var testutils = require('../utils/index');
      var options = {
          filename: './spec/integration/data/test-issue-99.xlsx',
          useStyles: true,
          useSharedStrings: true
      };
      var wb = new Excel.stream.xlsx.WorkbookWriter(options);
      var sheet = wb.addWorksheet("test");
      sheet.commit();
      wb.commit();

      var wb2 = new Excel.Workbook();
      wb2.xlsx.readFile('./spec/integration/data/test-issue-99.xlsx');
      throw Error("check the file is correclty opened by excel");

  });
  it('issue 163 - Error while using xslx readFile method', function() {
    var wb = new Excel.Workbook();
    return wb.xlsx.readFile('./spec/integration/data/test-issue-163.xlsx')
      .then(function() {
        // arriving here is success
        expect(true).to.equal(true);
      });
  });
  it('issue 176 - Unexpected xml node in parseOpen', function() {
    var wb = new Excel.Workbook();
    return wb.xlsx.readFile('./spec/integration/data/test-issue-176.xlsx')
      .then(function() {
        // arriving here is success
        expect(true).to.equal(true);
      });
  });
  describe('issue 219 - 1904 dates not supported', function() {
    it('Reading 1904.xlsx', function () {
      var wb = new Excel.Workbook();
      return wb.xlsx.readFile('./spec/integration/data/1904.xlsx')
        .then(function () {
          expect(wb.properties.date1904).to.equal(true);

          var ws = wb.getWorksheet('Sheet1');
          expect(ws.getCell('B4').value.toISOString()).to.equal('1904-01-01T00:00:00.000Z');
        });
    });
    it('Writing and Reading', function() {
      var wb = new Excel.Workbook();
      wb.properties.date1904 = true;
      var ws = wb.addWorksheet('Sheet1');
      ws.getCell('B4').value = new Date('1904-01-01T00:00:00.000Z');
      return wb.xlsx.writeFile(TEST_XLSX_FILE_NAME)
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(function(wb2) {
          expect(wb2.properties.date1904).to.equal(true);

          var ws2 = wb2.getWorksheet('Sheet1');
          expect(ws2.getCell('B4').value.toISOString()).to.equal('1904-01-01T00:00:00.000Z');
        });
    });
  });
});