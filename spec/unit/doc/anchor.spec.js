var expect = require('chai').expect;
var Anchor = require('../../../lib/doc/anchor');
var createSheetMock = require('../../utils/index').createSheetMock;

describe('Anchor', function () {
  describe('colWidth', function () {
    it('should colWidth equals 640000 when worksheet is undefined', function () {
      var anchor = new Anchor();
      expect(anchor.colWidth).to.equal(640000)
    });
    it('should colWidth equals 640000 when column has not set custom width', function () {
      var anchor = new Anchor();
      anchor.worksheet = createSheetMock();
      expect(anchor.colWidth).to.equal(640000)
    });
    it('should colWidth equals column width', function () {
      var anchor = new Anchor();
      var worksheet = createSheetMock();
      worksheet.addColumn(anchor.nativeCol + 1, {
        width: 10
      });
      anchor.worksheet = worksheet;
      expect(anchor.colWidth).to.equal(worksheet.getColumn(anchor.nativeCol + 1).width * 10000)
    });
  });
  describe('rowHeight', function () {
    it('should rowHeight equals 180000 when worksheet is undefined', function () {
      var anchor = new Anchor();
      expect(anchor.rowHeight).to.equal(180000)
    });
    it('should rowHeight equals 180000 when row has not set height', function () {
      var anchor = new Anchor();
      anchor.worksheet = createSheetMock();
      expect(anchor.rowHeight).to.equal(180000)
    });
    it('should rowHeight equals row height', function () {
      var worksheet = createSheetMock();
      worksheet.getRow(1).height = 10;

      var anchor = new Anchor();
      anchor.worksheet = worksheet;
      expect(anchor.rowHeight).to.equal(worksheet.getRow(1).height * 10000);
    });
  });
  describe('resize worksheet`s cells', function () {
    var context = this;
    before(function () {
      context.worksheet = createSheetMock();
      context.worksheet.getColumn(1).width = 20;
      context.worksheet.getRow(1).height = 20;

      context.anchor = new Anchor({col: 0.6, row: 0.6});
      context.anchor.worksheet = context.worksheet;

    });

    it('should update colWidth', function () {
      var pre = context.anchor.colWidth;
      context.worksheet.getColumn(1).width *= 2;
      expect(context.anchor.colWidth).to.not.equal(pre);
      expect(context.anchor.colWidth).to.equal(pre * 2);
    });
    it('should update rowHeight', function () {
      var pre = context.anchor.rowHeight;
      context.worksheet.getRow(1).height *= 2;
      expect(context.anchor.rowHeight).to.not.equal(pre);
      expect(context.anchor.rowHeight).to.equal(pre * 2);
    });
    it('should recalculate col', function () {
      var pre = context.anchor.col;
      context.worksheet.getColumn(1).width *= 2;
      expect(context.anchor.col).to.not.equal(pre);
    });
    it('should recalculate row', function () {
      var pre = context.anchor.row;
      context.worksheet.getRow(1).height *= 2;
      expect(context.anchor.row).to.not.equal(pre);
    });
    it('should integer part of row and rowOff should be always equals', function () {
      expect(Math.floor(context.anchor.row)).to.equal(Math.floor(context.anchor.nativeCol));
      context.worksheet.getRow(1).height *= 2;
      expect(Math.floor(context.anchor.row)).to.equal(Math.floor(context.anchor.nativeCol));
      context.worksheet.getRow(1).height /= 4;
      expect(Math.floor(context.anchor.row)).to.equal(Math.floor(context.anchor.nativeCol));
      context.worksheet.getRow(1).height = 0.1;
      expect(Math.floor(context.anchor.row)).to.equal(Math.floor(context.anchor.nativeCol));
      context.worksheet.getRow(1).height = 9999;
      expect(Math.floor(context.anchor.row)).to.equal(Math.floor(context.anchor.nativeCol));
    });
    it('should integer part of col and colOff should be always equals', function () {
      expect(Math.floor(context.anchor.col)).to.equal(Math.floor(context.anchor.nativeCol));
      context.worksheet.getColumn(1).width *= 2;
      expect(Math.floor(context.anchor.col)).to.equal(Math.floor(context.anchor.nativeCol));
      context.worksheet.getColumn(1).width /= 4;
      expect(Math.floor(context.anchor.col)).to.equal(Math.floor(context.anchor.nativeCol));
      context.worksheet.getColumn(1).width = 0.1;
      expect(Math.floor(context.anchor.col)).to.equal(Math.floor(context.anchor.nativeCol));
      context.worksheet.getColumn(1).width = 9999;
      expect(Math.floor(context.anchor.col)).to.equal(Math.floor(context.anchor.nativeCol));
    });
    it('should update nativeColOff after col has been changed', function () {
      var pre = context.anchor.nativeColOff;
      context.anchor.col -= 0.321;
      expect(context.anchor.nativeColOff).to.not.equal(pre);
    });
    it('should update nativeRowOff after row has been changed', function () {
      var pre = context.anchor.nativeRowOff;
      context.anchor.row -= 0.321;
      expect(context.anchor.nativeColOff).to.not.equal(pre);
    });
  });
});
