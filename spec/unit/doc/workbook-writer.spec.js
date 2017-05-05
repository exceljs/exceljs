'use strict';

var Stream = require('stream');

var expect = require('chai').expect;

var Excel = require('../../../excel');

describe('Workbook Writer', function() {
  it('returns undefined for non-existant sheet', function() {
    var stream = new Stream.Writable({write: function noop() {}});
    var wb = new Excel.stream.xlsx.WorkbookWriter({
      stream: stream
    });
    wb.addWorksheet('first');
    expect(wb.getWorksheet('w00t')).to.equal(undefined);
  });
});
