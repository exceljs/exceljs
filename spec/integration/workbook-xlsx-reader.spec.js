'use strict';

//var expect = require('chai').expect;
var Bluebird = require('bluebird');
var fs = require('fs');
var fsa = Bluebird.promisifyAll(fs);
//var _ = require('underscore');
var Excel = require('../../excel');
var testutils = require('./../testutils');
//var utils = require('../../lib/utils/utils');

// need some architectural changes to make stream read work properly
// because of: shared strings, sheet names, etc are not read in guaranteed order
describe.skip('WorkbookReader', function() {
  describe('Serialise', function() {
    after(function() {
      // delete the working file, don't care about errors
      return fsa.unlinkAsync('./wbr.test.xlsx').catch(function(){});
    });

    it('xlsx file', function() {
      var wb = testutils.createTestBook(true, Excel.Workbook);

      return wb.xlsx.writeFile('./wbr.test.xlsx')
        .then(function() {
          return testutils.checkTestBookReader('./wbr.test.xlsx');
        });
    });
  });
});
