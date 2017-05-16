var expect = require('chai').expect;

var Excel = require('../../../excel');
var Dimensions = require('../../../lib/doc/range');
var testUtils = require('../../utils/index');

describe('Worksheet', function() {
  describe('Merge Cells', function() {
    it('references the same top-left value', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      // initial values
      ws.getCell('A1').value = 'A1';
      ws.getCell('B1').value = 'B1';
      ws.getCell('A2').value = 'A2';
      ws.getCell('B2').value = 'B2';

      ws.mergeCells('A1:B2');

      expect(ws.getCell('A1').value).to.equal('A1');
      expect(ws.getCell('B1').value).to.equal('A1');
      expect(ws.getCell('A2').value).to.equal('A1');
      expect(ws.getCell('B2').value).to.equal('A1');

      expect(ws.getCell('A1').type).to.equal(Excel.ValueType.String);
      expect(ws.getCell('B1').type).to.equal(Excel.ValueType.Merge);
      expect(ws.getCell('A2').type).to.equal(Excel.ValueType.Merge);
      expect(ws.getCell('B2').type).to.equal(Excel.ValueType.Merge);
    });

    it('does not allow overlapping merges', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      ws.mergeCells('B2:C3');

      // intersect four corners
      expect(function() { ws.mergeCells('A1:B2'); }).to.throw(Error);
      expect(function() { ws.mergeCells('C1:D2'); }).to.throw(Error);
      expect(function() { ws.mergeCells('C3:D4'); }).to.throw(Error);
      expect(function() { ws.mergeCells('A3:B4'); }).to.throw(Error);

      // enclosing
      expect(function() { ws.mergeCells('A1:D4'); }).to.throw(Error);
    });

    it('merges and unmerges', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      var expectMaster = function(range, master) {
        var d = new Dimensions(range);
        for (var i = d.top; i <= d.bottom; i++) {
          for (var j = d.left; j <= d.right; j++) {
            var cell = ws.getCell(i, j);
            var masterCell = master ? ws.getCell(master) : cell;
            expect(cell.master.address).to.equal(masterCell.address);
          }
        }
      };

      // merge some cells, then unmerge them
      ws.mergeCells('A1:B2');
      expectMaster('A1:B2', 'A1');
      ws.unMergeCells('A1:B2');
      expectMaster('A1:B2', null);

      // unmerge just one cell
      ws.mergeCells('A1:B2');
      expectMaster('A1:B2', 'A1');
      ws.unMergeCells('A1');
      expectMaster('A1:B2', null);

      ws.mergeCells('A1:B2');
      expectMaster('A1:B2', 'A1');
      ws.unMergeCells('B2');
      expectMaster('A1:B2', null);

      // build 4 merge-squares
      ws.mergeCells('A1:B2');
      ws.mergeCells('D1:E2');
      ws.mergeCells('A4:B5');
      ws.mergeCells('D4:E5');

      expectMaster('A1:B2', 'A1');
      expectMaster('D1:E2', 'D1');
      expectMaster('A4:B5', 'A4');
      expectMaster('D4:E5', 'D4');

      // unmerge the middle
      ws.unMergeCells('B2:D4');

      expectMaster('A1:B2', null);
      expectMaster('D1:E2', null);
      expectMaster('A4:B5', null);
      expectMaster('D4:E5', null);
    });

    it('does not allow overlapping merges', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      ws.mergeCells('B2:C3');

      // intersect four corners
      expect(function() { ws.mergeCells('A1:B2'); }).to.throw(Error);
      expect(function() { ws.mergeCells('C1:D2'); }).to.throw(Error);
      expect(function() { ws.mergeCells('C3:D4'); }).to.throw(Error);
      expect(function() { ws.mergeCells('A3:B4'); }).to.throw(Error);

      // enclosing
      expect(function() { ws.mergeCells('A1:D4'); }).to.throw(Error);
    });

    it('merges styles', function() {
      var wb = new Excel.Workbook();
      var ws = wb.addWorksheet('blort');

      // initial value
      var B2 = ws.getCell('B2');
      B2.value = 5;
      B2.style.font = testUtils.styles.fonts.broadwayRedOutline20;
      B2.style.border = testUtils.styles.borders.doubleRed;
      B2.style.fill = testUtils.styles.fills.blueWhiteHGrad;
      B2.style.alignment = testUtils.styles.namedAlignments.middleCentre;
      B2.style.numFmt = testUtils.styles.numFmts.numFmt1;

      // expecting styles to be copied (see worksheet spec)
      ws.mergeCells('B2:C3');

      expect(ws.getCell('B2').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
      expect(ws.getCell('B2').border).to.deep.equal(testUtils.styles.borders.doubleRed);
      expect(ws.getCell('B2').fill).to.deep.equal(testUtils.styles.fills.blueWhiteHGrad);
      expect(ws.getCell('B2').alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);
      expect(ws.getCell('B2').numFmt).to.deep.equal(testUtils.styles.numFmts.numFmt1);

      expect(ws.getCell('B3').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
      expect(ws.getCell('B3').border).to.deep.equal(testUtils.styles.borders.doubleRed);
      expect(ws.getCell('B3').fill).to.deep.equal(testUtils.styles.fills.blueWhiteHGrad);
      expect(ws.getCell('B3').alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);
      expect(ws.getCell('B3').numFmt).to.deep.equal(testUtils.styles.numFmts.numFmt1);

      expect(ws.getCell('C2').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
      expect(ws.getCell('C2').border).to.deep.equal(testUtils.styles.borders.doubleRed);
      expect(ws.getCell('C2').fill).to.deep.equal(testUtils.styles.fills.blueWhiteHGrad);
      expect(ws.getCell('C2').alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);
      expect(ws.getCell('C2').numFmt).to.deep.equal(testUtils.styles.numFmts.numFmt1);

      expect(ws.getCell('C3').font).to.deep.equal(testUtils.styles.fonts.broadwayRedOutline20);
      expect(ws.getCell('C3').border).to.deep.equal(testUtils.styles.borders.doubleRed);
      expect(ws.getCell('C3').fill).to.deep.equal(testUtils.styles.fills.blueWhiteHGrad);
      expect(ws.getCell('C3').alignment).to.deep.equal(testUtils.styles.namedAlignments.middleCentre);
      expect(ws.getCell('C3').numFmt).to.deep.equal(testUtils.styles.numFmts.numFmt1);
    });
  });
});
