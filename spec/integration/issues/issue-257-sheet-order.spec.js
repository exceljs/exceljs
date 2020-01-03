'use strict';

var chai = require('chai');
var verquire = require('../../utils/verquire');

var Excel = verquire('excel');

var expect = chai.expect;

describe('github issues', function() {
  it('issue 257 - worksheet order is not respected', function() {
    var wb = new Excel.Workbook();
    return wb.xlsx.readFile('./spec/integration/data/test-issue-257.xlsx')
      .then(function() {
        expect(wb.worksheets.map(ws => ws.name)).to.deep.equal(['First', 'Second']);
      });
  });
});
