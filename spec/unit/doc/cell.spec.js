'use strict';

var expect = require('chai').expect;

var colCache = require('../../../lib/utils/col-cache');
var Cell = require('../../../lib/doc/cell');
var Enums = require('../../../lib/doc/enums');

var sheetMock = {
  reset() {
    this.rows = [];
    this.columns = [];
  },
  findRow(num) {
    return this.rows[num];
  },
  getRow(num) {
    return this.rows[num] ||
      this.createRow(num);
  },
  findColumn(num) {
    return this.columns[num];
  },
  getColumn(num) {
    return this.columns[num] ||
      this.createColumn(num);
  },
  createRow(num) {
    this.rows[num] = {
      cells: [],
      findCell(col) {
        return this.cells[col];
      },
      getCell(col) {
        return this.cells[col] ||
          this.createCell(col);
      },
      createCell(col) {
        var address = colCache.encodeAddress(this.number, col);
        var column = sheetMock.getColumn(col);
        return (this.cells[col] = new Cell(this, column, address));
      },
      number: num,
      get worksheet() {
        return sheetMock;
      },
    };
    return this.rows[num];
  },
  createColumn: function(num) {
    this.columns[num] = {
      number: num,
      letter: colCache.n2l(num),
    };
    return this.columns[num];
  },
  getCell: function(address) {
    var fullAddress = colCache.decodeAddress(address);
    var row = this.getRow(fullAddress.row);
    return row.getCell(fullAddress.col);
  },
  findCell: function(address) {
    var fullAddress = colCache.decodeAddress(address);
    var row = this.getRow(fullAddress.row);
    return row && row.findCell(fullAddress.col);
  },
};
var fonts = {
  arialBlackUI14: { name: 'Arial Black', family: 2, size: 14, underline: true, italic: true },
  comicSansUdB16: { name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true },
  broadwayRedOutline20: { name: 'Broadway', family: 5, size: 20, outline: true, color: { argb: 'FFFF0000'}}
};

describe('Cell', function() {
  beforeEach(() => {
    sheetMock.reset();
  });
  it('stores values', function() {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    var a1 = sheetMock.getCell('A1');

    expect(a1.address).to.equal('A1');
    expect(a1.$col$row).to.equal('$A$1');

    expect(a1.type).to.equal(Enums.ValueType.Null);

    expect(a1.value = 5).to.equal(5);
    expect(a1.value).to.equal(5);
    expect(a1.type).to.equal(Enums.ValueType.Number);

    var strValue = 'Hello, World!';
    expect(a1.value = strValue).to.equal(strValue);
    expect(a1.value).to.equal(strValue);
    expect(a1.type).to.equal(Enums.ValueType.String);

    var dateValue = new Date();
    expect(a1.value = dateValue).to.equal(dateValue);
    expect(a1.value).to.equal(dateValue);
    expect(a1.type).to.equal(Enums.ValueType.Date);

    var formulaValue = {formula: 'A2', result: 5};
    expect(a1.value = formulaValue).to.deep.equal(formulaValue);
    expect(a1.value).to.deep.equal(formulaValue);
    expect(a1.type).to.equal(Enums.ValueType.Formula);

    // no result
    formulaValue = {formula: 'A3'};
    expect(a1.value = formulaValue).to.deep.equal(formulaValue);
    expect(a1.value).to.deep.equal({formula: 'A3', result: undefined});
    expect(a1.type).to.equal(Enums.ValueType.Formula);

    var hyperlinkValue = {hyperlink: 'http://www.link.com', text: 'www.link.com'};
    expect(a1.value = hyperlinkValue).to.deep.equal(hyperlinkValue);
    expect(a1.value).to.deep.equal(hyperlinkValue);
    expect(a1.type).to.equal(Enums.ValueType.Hyperlink);

    expect(a1.value = null).to.be.null();
    expect(a1.type).to.equal(Enums.ValueType.Null);

    expect(a1.value = {json: 'data'}).to.deep.equal({json: 'data'});
    expect(a1.type).to.equal(Enums.ValueType.String);
  });
  it('validates options on construction', function() {
    var row = sheetMock.getRow(1);
    var column = sheetMock.getColumn(1);

    expect(function() { new Cell(); }).to.throw(Error);
    expect(function() { new Cell(row); }).to.throw(Error);
    expect(function() { new Cell(row, 'A'); }).to.throw(Error);
    expect(function() { new Cell(row, 'Hello, World!'); }).to.throw(Error);
    expect(function() { new Cell(null, null, 'A1'); }).to.throw(Error);
    expect(function() { new Cell(row, null, 'A1'); }).to.throw(Error);
    expect(function() { new Cell(null, column, 'A1'); }).to.throw(Error);
  });
  it('merges', function() {
    var a1 = sheetMock.getCell('A1');
    var a2 = sheetMock.getCell('A2');

    a1.value = 5;
    a2.value = 'Hello, World!';

    a2.merge(a1);

    expect(a2.value).to.equal(5);
    expect(a2.type).to.equal(Enums.ValueType.Merge);
    expect(a1._mergeCount).to.equal(1);
    expect(a1.isMerged).to.be.ok();
    expect(a2.isMerged).to.be.ok();
    expect(a2.isMergedTo(a1)).to.be.ok();
    expect(a2.master).to.equal(a1);
    expect(a1.master).to.equal(a1);

    // assignment of slaves write to the master
    a2.value = 7;
    expect(a1.value).to.equal(7);

    // assignment of strings should add 1 ref
    var strValue = 'Boo!';
    a2.value = strValue;
    expect(a1.value).to.equal(strValue);

    // unmerge should work also
    a2.unmerge();
    expect(a2.type).to.equal(Enums.ValueType.Null);
    expect(a1._mergeCount).to.equal(0);
    expect(a1.isMerged).to.not.be.ok();
    expect(a2.isMerged).to.not.be.ok();
    expect(a2.isMergedTo(a1)).to.not.be.ok();
    expect(a2.master).to.equal(a2);
    expect(a1.master).to.equal(a1);
  });

  it('upgrades from string to hyperlink', function() {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    var a1 = sheetMock.getCell('A1');

    var strValue = 'www.link.com';
    var linkValue = 'http://www.link.com';

    a1.value = strValue;

    a1._upgradeToHyperlink(linkValue);

    expect(a1.type).to.equal(Enums.ValueType.Hyperlink);
  });

  it('doesn\'t upgrade from non-string to hyperlink', function() {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    var a1 = sheetMock.getCell('A1');

    var linkValue = 'http://www.link.com';

    // null
    a1._upgradeToHyperlink(linkValue);
    expect(a1.type).to.equal(Enums.ValueType.Null);

    // number
    a1.value = 5;
    a1._upgradeToHyperlink(linkValue);
    expect(a1.type).to.equal(Enums.ValueType.Number);

    // date
    a1.value = new Date();
    a1._upgradeToHyperlink(linkValue);
    expect(a1.type).to.equal(Enums.ValueType.Date);

    // formula
    a1.value = {formula: 'A2'};
    a1._upgradeToHyperlink(linkValue);
    expect(a1.type).to.equal(Enums.ValueType.Formula);

    // hyperlink
    a1.value = {hyperlink: 'http://www.link2.com', text: 'www.link2.com'};
    a1._upgradeToHyperlink(linkValue);
    expect(a1.type).to.deep.equal(Enums.ValueType.Hyperlink);

    // cleanup
    a1.value = null;
  });

  it('inherits column styles', function() {
    sheetMock.getRow(1);
    var column = sheetMock.getColumn(1);

    column.style = {
      font: fonts.arialBlackUI14
    };

    var a1 = sheetMock.getCell('A1');
    expect(a1.font).to.deep.equal(fonts.arialBlackUI14);
  });

  it('inherits row styles', function() {
    var row = sheetMock.getRow(1);
    sheetMock.getColumn(1);

    row.style = {
      font: fonts.broadwayRedOutline20
    };

    var a1 = sheetMock.getCell('A1');
    expect(a1.font).to.deep.equal(fonts.broadwayRedOutline20);
  });

  it('has effective types', function() {
    sheetMock.getRow(1);
    sheetMock.getColumn(1);

    var a1 = sheetMock.getCell('A1');

    expect(a1.type).to.equal(Enums.ValueType.Null);
    expect(a1.effectiveType).to.equal(Enums.ValueType.Null);

    a1.value = 5;
    expect(a1.type).to.equal(Enums.ValueType.Number);
    expect(a1.effectiveType).to.equal(Enums.ValueType.Number);

    a1.value = 'Hello, World!';
    expect(a1.type).to.equal(Enums.ValueType.String);
    expect(a1.effectiveType).to.equal(Enums.ValueType.String);

    a1.value = new Date();
    expect(a1.type).to.equal(Enums.ValueType.Date);
    expect(a1.effectiveType).to.equal(Enums.ValueType.Date);

    a1.value = {formula: 'A2', result: 5};
    expect(a1.type).to.deep.equal(Enums.ValueType.Formula);
    expect(a1.effectiveType).to.equal(Enums.ValueType.Number);

    a1.value = {formula: 'A2', result: 'Hello, World!'};
    expect(a1.type).to.deep.equal(Enums.ValueType.Formula);
    expect(a1.effectiveType).to.equal(Enums.ValueType.String);

    a1.value = {hyperlink: 'http://www.link.com', text: 'www.link.com'};
    expect(a1.type).to.deep.equal(Enums.ValueType.Hyperlink);
    expect(a1.effectiveType).to.equal(Enums.ValueType.Hyperlink);
  });

  it('shares formulas', () => {
    var a1 = sheetMock.getCell('A1');
    var b1 = sheetMock.getCell('B1');
    var c1 = sheetMock.getCell('C1');

    a1.value = 1;
    b1.value = { formula: 'A1+1', result: 2 };
    c1.value = { sharedFormula: 'B1', result: 3 };

    expect(b1.type).to.equal(Enums.ValueType.Formula);
    expect(b1.formulaType).to.equal(Enums.FormulaType.Master);
    expect(c1.type).to.equal(Enums.ValueType.Formula);
    expect(c1.formulaType).to.equal(Enums.FormulaType.Shared);
    expect(c1.formula).to.equal('B1+1');
  });

  it('escapes dangerous html', () => {
    var a1 = sheetMock.getCell('A1');

    a1.value = '<script>alert("yoohoo")</script>';

    expect(a1.html).to.equal('&lt;script&gt;alert(&quot;yoohoo&quot;)&lt;/script&gt;');
  });
});
