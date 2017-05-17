'use strict';

var chai = require('chai');

var verquire = require('../../utils/verquire');

var PromishLib = verquire('utils/promish');
var Excel = verquire('excel');

var expect = chai.expect;

// this file to contain integration tests created from github issues
var TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';

describe('github issues', function() {
  describe('issue 266 - Breaking change removing bluebird', function() {
    var promish;
    beforeEach(function() {
      promish = PromishLib.Promish;
    });
    afterEach(function() {
      PromishLib.Promish = promish;
    });

    it('common bluebird functions', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('Sheet1');
      var calledFinally = false;
      ws.getCell('A1').value = 'Hello, World!';
      return wb.xlsx.writeFile(TEST_XLSX_FILE_NAME)
        .then(function() {
          var wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(function() {
          return Promise.all([
            Promise.resolve('a'),
            Promise.resolve('b')
          ]);
        })
        .map(function(value) {
          return value + value;
        })
        .spread(function(a, b) {
          expect(a).to.equal('aa');
          expect(b).to.equal('bb');
          return 'c';
        })
        .finally(function() {
          calledFinally = true;
        })
        .then(function(value) {
          expect(value).to.equal('c');
          expect(calledFinally).to.equal(true);
          return Promise.all([
            Promise.resolve('a'),
            Promise.resolve('b')
          ]);
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
});
