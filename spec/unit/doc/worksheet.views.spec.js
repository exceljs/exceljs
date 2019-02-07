var expect = require('chai').expect;

var Excel = require('../../../excel');

describe('Worksheet', function() {
  describe('Views', function() {
      it('adjusts collapsed property of columns', function() {
        var wb = new Excel.Workbook();
        var ws = wb.addWorksheet('sheet1');

        var col1 = ws.getColumn(1);
        var col2 = ws.getColumn(2);
        var col3 = ws.getColumn(3);
        expect(col1.collapsed).to.be.false();
        expect(col2.collapsed).to.be.false();
        expect(col3.collapsed).to.be.false();

        col1.outlineLevel = 0;
        col2.outlineLevel = 1;
        col3.outlineLevel = 2;
        expect(col1.collapsed).to.be.false();
        expect(col2.collapsed).to.be.true();
        expect(col3.collapsed).to.be.true();

        ws.properties.outlineLevelCol = 2;
        expect(col1.collapsed).to.be.false();
        expect(col2.collapsed).to.be.false();
        expect(col3.collapsed).to.be.true();

        ws.properties.outlineLevelCol = 3;
        expect(col1.collapsed).to.be.false();
        expect(col2.collapsed).to.be.false();
        expect(col3.collapsed).to.be.false();
      });

      it('adjusts collapsed property of row', function() {
        var wb = new Excel.Workbook();
        var ws = wb.addWorksheet('sheet1');

        var row1 = ws.getRow(1);
        var row2 = ws.getRow(2);
        var row3 = ws.getRow(3);
        expect(row1.collapsed).to.be.false();
        expect(row2.collapsed).to.be.false();
        expect(row3.collapsed).to.be.false();

        row1.outlineLevel = 0;
        row2.outlineLevel = 1;
        row3.outlineLevel = 2;
        expect(row1.collapsed).to.be.false();
        expect(row2.collapsed).to.be.true();
        expect(row3.collapsed).to.be.true();

        ws.properties.outlineLevelRow = 2;
        expect(row1.collapsed).to.be.false();
        expect(row2.collapsed).to.be.false();
        expect(row3.collapsed).to.be.true();

        ws.properties.outlineLevelRow = 3;
        expect(row1.collapsed).to.be.false();
        expect(row2.collapsed).to.be.false();
        expect(row3.collapsed).to.be.false();
      });

      it('sets outline levels via column headers', function() {
        var wb = new Excel.Workbook();
        var ws = wb.addWorksheet('sheet1');

        ws.columns = [
          { key: 'id', width: 10, outlineLevel: 1 },
          { key: 'name', width: 32, outlineLevel: 2 },
          { key: 'dob', width: 10, outlineLevel: 3 }
        ];

        expect(ws.getColumn(1).outlineLevel).to.equal(1);
        expect(ws.getColumn(2).outlineLevel).to.equal(2);
        expect(ws.getColumn(3).outlineLevel).to.equal(3);
      });
  });
});
