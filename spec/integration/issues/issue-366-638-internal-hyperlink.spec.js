'use strict';

var chai = require('chai');

var verquire = require('../../utils/verquire');

var Excel = verquire('excel');

var expect = chai.expect;

var Enums = require('../../../lib/doc/enums');

describe('github issues', function() {
  it('issue 366 and 638 - Hyperlinks', function() {
    var wb = new Excel.Workbook();
    return wb.xlsx.readFile('./spec/integration/data/test-issue-366-638.xlsx')
      .then(function() {
        var sheet1 = wb.getWorksheet(1);
        const cell1 = sheet1.getCell('A1');
        expect(sheet1.getCell('A1').hyperlink.target).to.equal("Sheet2!A1");

        // set hyperlink on a formula
        var hyperlink = {mode: 'external', target: 'www.link.com'};
        var a3 = sheet1.getCell('A3');
        var formulaValue = a3.value = {formula: 'A4', result: ''};
        a3.hyperlink = hyperlink;
        expect(a3.hyperlink).to.deep.equal(hyperlink);
        expect(a3.value).to.deep.equal(formulaValue);
        expect(a3.type).to.equal(Enums.ValueType.Formula);


        wb.xlsx.writeFile('./spec/out/temp.xlsx');
      });

  });
});


