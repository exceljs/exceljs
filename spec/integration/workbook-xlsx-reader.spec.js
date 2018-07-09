'use strict';

var verquire = require('../utils/verquire');
var testutils = require('../utils/index');
var fs = require('fs');
var expect = require('chai').expect;

var Excel = verquire('excel');

var TEST_FILE_NAME = './spec/out/wb.test.xlsx';

// need some architectural changes to make stream read work properly
// because of: shared strings, sheet names, etc are not read in guaranteed order
describe('WorkbookReader', function() {
  describe('Serialise', function() {
    it('xlsx file', function() {
      this.timeout(10000);
      var wb = testutils.createTestBook(new Excel.Workbook(), 'xlsx');

      return wb.xlsx.writeFile(TEST_FILE_NAME)
        .then(function() {
          return testutils.checkTestBookReader(TEST_FILE_NAME);
        });
    });
  });

  describe('#readFile', function() {
    describe('Row limit', function() {
      it('should bail out if the file contains more rows than the limit', function() {
        var workbook = new Excel.Workbook();
        // The Fibonacci sheet has 19 rows
        return workbook.xlsx.readFile('./spec/integration/data/fibonacci.xlsx', {maxRows: 10})
          .then(function() {
            throw new Error('Promise unexpectedly fulfilled');
          }, function(err) {
            expect(err.message).to.equal('Max row count exceeded');
          });
      });

      it('should fail fast on a huge file', function() {
        this.timeout(20000);
        var workbook = new Excel.Workbook();
        return workbook.xlsx.readFile('./spec/integration/data/huge.xlsx', {maxRows: 100})
          .then(function() {
            throw new Error('Promise unexpectedly fulfilled');
          }, function(err) {
            expect(err.message).to.equal('Max row count exceeded');
          });
      });

      it('should parse fine if the limit is not exceeded', function() {
        var workbook = new Excel.Workbook();
        return workbook.xlsx.readFile('./spec/integration/data/fibonacci.xlsx', {maxRows: 20});
      });
    });
  });

  describe('#read', function() {
    describe('Row limit', function() {
      it('should bail out if the file contains more rows than the limit', function() {
        var workbook = new Excel.Workbook();
        // The Fibonacci sheet has 19 rows
        return workbook.xlsx.read(fs.createReadStream('./spec/integration/data/fibonacci.xlsx'), {maxRows: 10})
          .then(function() {
            throw new Error('Promise unexpectedly fulfilled');
          }, function(err) {
            expect(err.message).to.equal('Max row count exceeded');
          });
      });

      it('should parse fine if the limit is not exceeded', function() {
        var workbook = new Excel.Workbook();
        return workbook.xlsx.read(fs.createReadStream('./spec/integration/data/fibonacci.xlsx'), {maxRows: 20});
      });
    });
  });

  describe('with a spreadsheet that contains formulas', function() {
    before(function() {
      var testContext = this;
      var workbook = new Excel.Workbook();
      return workbook.xlsx.read(fs.createReadStream('./spec/integration/data/formulas.xlsx'))
        .then(function() {
          testContext.worksheet = workbook.getWorksheet();
        });
    });

    describe('with a cell that contains a regular formula', function() {
      beforeEach(function() {
        this.cell = this.worksheet.getCell('A2');
      });

      it('should be classified as a formula cell', function() {
        expect(this.cell.type).to.equal(Excel.ValueType.Formula);
        expect(this.cell.isFormula).to.be.true;
      });

      it('should have text corresponding to the evaluated formula result', function() {
        expect(this.cell.text).to.equal('someone@example.com');
      });

      it('should have the formula source', function() {
        expect(this.cell.model.formula).to.equal('_xlfn.CONCAT("someone","@example.com")');
      });
    });

    describe('with a cell that contains a hyperlinked formula', function() {
      beforeEach(function() {
        this.cell = this.worksheet.getCell('A1');
      });

      it('should be classified as a formula cell', function() {
        expect(this.cell.type).to.equal(Excel.ValueType.Hyperlink);
      });

      it('should have text corresponding to the evaluated formula result', function() {
        expect(this.cell.value.text).to.equal('someone@example.com');
      });

      it('should have the formula source', function() {
        expect(this.cell.model.formula).to.equal('_xlfn.CONCAT("someone","@example.com")');
      });

      it('should contain the linked url', function() {
        expect(this.cell.value.hyperlink).to.equal('mailto:someone@example.com');
        expect(this.cell.hyperlink).to.equal('mailto:someone@example.com');
      });
    });
  });

  describe('with a spreadsheet that contains a shared string with an escaped underscore', function() {
    before(function() {
      var testContext = this;
      var workbook = new Excel.Workbook();
      return workbook.xlsx.read(fs.createReadStream('./spec/integration/data/shared_string_with_escape.xlsx'))
        .then(function() {
          testContext.worksheet = workbook.getWorksheet();
        });
    });

    it('should decode the underscore', function() {
      const cell = this.worksheet.getCell('A1');
      expect(cell.value).to.equal('_x000D_');
    });
  });
});
