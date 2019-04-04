'use strict';

const chai = require('chai');

const verquire = require('../../utils/verquire');

const PromishLib = verquire('utils/promish');
const Excel = verquire('excel');

const expect = chai.expect;

// this file to contain integration tests created from github issues
const TEST_XLSX_FILE_NAME = './spec/out/wb.test.xlsx';

describe('github issues', () => {
  describe('issue 266 - Breaking change removing bluebird', () => {
    let promish;
    beforeEach(() => {
      promish = PromishLib.Promish;
    });
    afterEach(() => {
      PromishLib.Promish = promish;
    });

    it('common bluebird functions', () => {
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('Sheet1');
      let calledFinally = false;
      ws.getCell('A1').value = 'Hello, World!';
      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(() => Promise.all([Promise.resolve('a'), Promise.resolve('b')]))
        .map(value => value + value)
        .spread((a, b) => {
          expect(a).to.equal('aa');
          expect(b).to.equal('bb');
          return 'c';
        })
        .finally(() => {
          calledFinally = true;
        })
        .then(value => {
          expect(value).to.equal('c');
          expect(calledFinally).to.equal(true);
          return Promise.all([Promise.resolve('a'), Promise.resolve('b')]);
        });
    });

    it('Promise dependency injection', () => {
      // to test that a promise implementation can be injected
      const Bluebird = require('bluebird');
      Excel.config.setValue('promise', Bluebird);
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Hello, World!';

      const promise = wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
      const bb = new Bluebird(() => {});

      expect(promise.then).to.equal(bb.then);
      return promise;
    });

    it('read and write with bluebird', () => {
      // to test that the bluebird is compatible with reading and writing
      const Bluebird = require('bluebird');
      Excel.config.setValue('promise', Bluebird);
      const wb = new Excel.Workbook();
      const ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Hello, World!';

      return wb.xlsx
        .writeFile(TEST_XLSX_FILE_NAME)
        .then(() => {
          const wb2 = new Excel.Workbook();
          return wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        })
        .then(wb2 => {
          const ws2 = wb2.getWorksheet('Sheet1');
          expect(ws2.getCell('A1').value).to.equal('Hello, World!');
        });
    });
  });
});
