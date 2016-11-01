'use strict';

var chai = require('chai');
var expect = chai.expect;
chai.use(require('chai-datetime'));

var Excel = require('../../excel');

// this file to contain integration tests created from github issues

describe('github issues', function() {
  it('issue 163 - Error while using xslx readFile method', function() {
    var wb = new Excel.Workbook();
    return wb.xlsx.readFile('./spec/integration/data/test-issue-163.xlsx')
      .then(function() {
        // arriving here is success
        expect(true).to.equal(true);
      });
  });
  it.only('issue 176 - Unexpected xml node in parseOpen', function() {
    var wb = new Excel.Workbook();
    return wb.xlsx.readFile('./spec/integration/data/test-issue-176.xlsx')
      .then(function() {
        // arriving here is success
        expect(true).to.equal(true);
      });
  });
});