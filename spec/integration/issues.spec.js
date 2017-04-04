'use strict';

var chai = require('chai');
var Excel = require('../../excel');
var PromishLib = require('../../lib/utils/promish');

var expect = chai.expect;
chai.use(require('chai-datetime'));

// this file to contain integration tests created from github issues
var TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';


describe('github issues', function() {
  it('issue 275 - hyperlink with query arguments corrupts workbook', function() {
    var options = {
      filename: TEST_XLSX_FILE_NAME,
      useStyles: true
    };
    var wb = new Excel.stream.xlsx.WorkbookWriter(options);
    var ws = wb.addWorksheet('Sheet1');

    var hyperlink = {
      text:'Somewhere with query params',
      hyperlink: 'www.somewhere.com?a=1&b=2&c=<>&d="\'"'
    };

    // Start of Heading
    ws.getCell('A1').value = hyperlink;
    ws.commit();

    return wb.commit()
      .then(function() {
        var wb2 = new Excel.Workbook();
        return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      })
      .then(function(wb2) {
        var ws2 = wb2.getWorksheet('Sheet1');
        expect(ws2.getCell('A1').value).to.deep.equal(hyperlink);
      });
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

  it('issue 234 - Broken XLSX because of "vertical tab" ascii character in a cell', function() {
    var wb = new Excel.Workbook();
    var ws = wb.addWorksheet('Sheet1');

    // Start of Heading
    ws.getCell('A1').value = 'Hello, \x01World!';

    // Vertical Tab
    ws.getCell('A2').value = 'Hello, \x0bWorld!';

    return wb.xlsx.writeFile(TEST_XLSX_FILE_NAME)
      .then(function() {
        var wb2 = new Excel.Workbook();
        return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
      })
      .then(function(wb2) {
        var ws2 = wb2.getWorksheet('Sheet1');
        expect(ws2.getCell('A1').value).to.equal('Hello, World!');
        expect(ws2.getCell('A2').value).to.equal('Hello, World!');
      });
  });

  describe('issue 266 - Breaking change removing bluebird', function() {
    var promish;
    beforeEach(function () {
      promish = PromishLib.Promish;
    });
    afterEach(function () {
      PromishLib.Promish = promish;
    });

    it('common bluebird functions', function () {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('Sheet1');
      var calledFinally = false;
      ws.getCell('A1').value = 'Hello, World!';
      return wb.xlsx.writeFile(TEST_XLSX_FILE_NAME)
        .then(function () {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(function () {
          return Promise.all([
            Promise.resolve('a'),
            Promise.resolve('b')
          ])
        })
        .map(function (value) {
          return value + value;
        })
        .spread(function (a, b) {
          expect(a).to.equal('aa');
          expect(b).to.equal('bb');
          return 'c';
        })
        .finally(function () {
          calledFinally = true;
        })
        .then(function (value) {
          expect(value).to.equal('c');
          expect(calledFinally).to.equal(true);
          return Promise.all([
            Promise.resolve('a'),
            Promise.resolve('b')
          ])
        });
    });

    it('Promise dependency injection', function() {
      // to test that a promise implementation can be injected
      var Bluebird = require('bluebird');
      Excel.config.setValue('promise', Bluebird);
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Hello, World!';

      var promise = wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      var bb = new Bluebird(function() {});

      expect(promise.then).to.equal(bb.then);
      return promise;
    });

    it('read and write with bluebird', function() {
      // to test that the bluebird is compatible with reading and writing
      var Bluebird = require('bluebird');
      Excel.config.setValue('promise', Bluebird);
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Hello, World!';

      return wb.xlsx.writeFile(TEST_XLSX_FILE_NAME)
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(function(wb2) {
          var ws2 = wb2.getWorksheet('Sheet1');
          expect(ws2.getCell('A1').value).to.equal('Hello, World!');
        });
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