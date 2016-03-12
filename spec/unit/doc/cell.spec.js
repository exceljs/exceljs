'use strict';

var expect = require('chai').expect;

var Cell = require('../../../lib/doc/cell');
var Enums = require('../../../lib/doc/enums');

var mock = {
  row: function() {
    return {
      cells: {},
      getCell: function(address) {
        return this.cells[address];
      },
      number: 1
    };
  },
  column: function() {
    return {
      number: 1,
      letter: 'A'
    };
  }
};
var fonts = {
  arialBlackUI14: { name: 'Arial Black', family: 2, size: 14, underline: true, italic: true },
  comicSansUdB16: { name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true },
  broadwayRedOutline20: { name: 'Broadway', family: 5, size: 20, outline: true, color: { argb:'FFFF0000'}}
};

describe('Cell', function() {
  it('stores values', function() {
    var row = mock.row();
    var column = mock.column();

    row.cells.A1 = new Cell(row, column, 'A1');

    expect(row.cells.A1.address).to.equal('A1');
    expect(row.cells.A1.$col$row).to.equal('$A$1');

    expect(row.cells.A1.type).to.equal(Enums.ValueType.Null);

    expect(row.cells.A1.value = 5).to.equal(5);
    expect(row.cells.A1.value).to.equal(5);
    expect(row.cells.A1.type).to.equal(Enums.ValueType.Number);

    var strValue = 'Hello, World!';
    expect(row.cells.A1.value = strValue).to.equal(strValue);
    expect(row.cells.A1.value).to.equal(strValue);
    expect(row.cells.A1.type).to.equal(Enums.ValueType.String);

    var dateValue = new Date();
    expect(row.cells.A1.value = dateValue).to.equal(dateValue);
    expect(row.cells.A1.value).to.equal(dateValue);
    expect(row.cells.A1.type).to.equal(Enums.ValueType.Date);

    var formulaValue = {formula: 'A2', result: 5};
    expect(row.cells.A1.value = formulaValue).to.deep.equal(formulaValue);
    expect(row.cells.A1.value).to.deep.equal(formulaValue);
    expect(row.cells.A1.type).to.equal(Enums.ValueType.Formula);

    // no result
    formulaValue = {formula: 'A3'};
    expect(row.cells.A1.value = formulaValue).to.deep.equal(formulaValue);
    expect(row.cells.A1.value).to.deep.equal({formula: 'A3', result: undefined});
    expect(row.cells.A1.type).to.equal(Enums.ValueType.Formula);

    var hyperlinkValue = {hyperlink: 'http://www.link.com', text: 'www.link.com'};
    expect(row.cells.A1.value = hyperlinkValue).to.deep.equal(hyperlinkValue);
    expect(row.cells.A1.value).to.deep.equal(hyperlinkValue);
    expect(row.cells.A1.type).to.equal(Enums.ValueType.Hyperlink);

    expect(row.cells.A1.value = null).to.be.null;
    expect(row.cells.A1.type).to.equal(Enums.ValueType.Null);
  });
  it('validates options on construction', function() {
    var row = mock.row();
    var column = mock.column();
    expect(function() { new Cell(); }).to.throw(Error);
    expect(function() { new Cell(row); }).to.throw(Error);
    expect(function() { new Cell(row, 'A'); }).to.throw(Error);
    expect(function() { new Cell(row, 'Hello, World!'); }).to.throw(Error);
    expect(function() { new Cell(null, null, 'A1'); }).to.throw(Error);
    expect(function() { new Cell(row, null, 'A1'); }).to.throw(Error);
    expect(function() { new Cell(null, column, 'A1'); }).to.throw(Error);
  });
  it('merges', function() {
    var row = mock.row();
    var column = mock.column();

    row.cells.A1 = new Cell(row, column, 'A1');
    row.cells.A2 = new Cell(row, column, 'A2');

    row.cells.A1.value = 5;
    row.cells.A2.value = 'Hello, World!';

    row.cells.A2.merge(row.cells.A1);

    expect(row.cells.A2.value).to.equal(5);
    expect(row.cells.A2.type).to.equal(Enums.ValueType.Merge);
    expect(row.cells.A1._mergeCount).to.equal(1);
    expect(row.cells.A1.isMerged).to.be.ok;
    expect(row.cells.A2.isMerged).to.be.ok;
    expect(row.cells.A2.isMergedTo(row.cells.A1)).to.be.ok;
    expect(row.cells.A2.master).to.equal(row.cells.A1);
    expect(row.cells.A1.master).to.equal(row.cells.A1);

    // assignment of slaves write to the master
    row.cells.A2.value = 7;
    expect(row.cells.A1.value).to.equal(7);

    // assignment of strings should add 1 ref
    var strValue = 'Boo!';
    row.cells.A2.value = strValue;
    expect(row.cells.A1.value).to.equal(strValue);

    // unmerge should work also
    row.cells.A2.unmerge();
    expect(row.cells.A2.type).to.equal(Enums.ValueType.Null);
    expect(row.cells.A1._mergeCount).to.equal(0);
    expect(row.cells.A1.isMerged).to.not.be.ok;
    expect(row.cells.A2.isMerged).to.not.be.ok;
    expect(row.cells.A2.isMergedTo(row.cells.A1)).to.not.be.ok;
    expect(row.cells.A2.master).to.equal(row.cells.A2);
    expect(row.cells.A1.master).to.equal(row.cells.A1);
  });

  it('upgrades from string to hyperlink', function() {
    var row = mock.row();
    var column = mock.column();

    var strValue = 'www.link.com';
    var linkValue = 'http://www.link.com';

    row.cells.A1 = new Cell(row, column, 'A1');
    row.cells.A1.value = strValue;

    row.cells.A1._upgradeToHyperlink(linkValue);

    expect(row.cells.A1.type).to.equal(Enums.ValueType.Hyperlink);
  });

  it('doesn\'t upgrade from non-string to hyperlink', function() {
    var row = mock.row();
    var column = mock.column();

    var linkValue = 'http://www.link.com';

    row.cells.A1 = new Cell(row, column, 'A1');

    // null
    row.cells.A1._upgradeToHyperlink(linkValue);
    expect(row.cells.A1.type).to.equal(Enums.ValueType.Null);

    // number
    row.cells.A1.value = 5;
    row.cells.A1._upgradeToHyperlink(linkValue);
    expect(row.cells.A1.type).to.equal(Enums.ValueType.Number);

    // date
    row.cells.A1.value = new Date();
    row.cells.A1._upgradeToHyperlink(linkValue);
    expect(row.cells.A1.type).to.equal(Enums.ValueType.Date);

    // formula
    row.cells.A1.value = {formula: 'A2'};
    row.cells.A1._upgradeToHyperlink(linkValue);
    expect(row.cells.A1.type).to.equal(Enums.ValueType.Formula);

    // hyperlink
    row.cells.A1.value = {hyperlink: 'http://www.link2.com', text: 'www.link2.com'};
    row.cells.A1._upgradeToHyperlink(linkValue);
    expect(row.cells.A1.type).to.deep.equal(Enums.ValueType.Hyperlink);

    // cleanup
    row.cells.A1.value = null;
  });

  it('inherits row and column styles', function() {
    var row = mock.row();
    var column = mock.column();

    column.style = {
      font: fonts.arialBlackUI14
    };

    row.cells.A1 = new Cell(row, column, 'A1');
    expect(row.cells.A1.font).to.deep.equal(fonts.arialBlackUI14);

    row.style = {
      font: fonts.broadwayRedOutline20
    };
    row.cells.A1 = new Cell(row, column, 'A1');
    expect(row.cells.A1.font).to.deep.equal(fonts.broadwayRedOutline20);
  });

  it('has effective types', function() {
    var row = mock.row();
    var column = mock.column();

    row.cells.A1 = new Cell(row, column, 'A1');
    expect(row.cells.A1.type).to.equal(Enums.ValueType.Null);
    expect(row.cells.A1.effectiveType).to.equal(Enums.ValueType.Null);

    row.cells.A1.value = 5;
    expect(row.cells.A1.type).to.equal(Enums.ValueType.Number);
    expect(row.cells.A1.effectiveType).to.equal(Enums.ValueType.Number);

    row.cells.A1.value = 'Hello, World!';
    expect(row.cells.A1.type).to.equal(Enums.ValueType.String);
    expect(row.cells.A1.effectiveType).to.equal(Enums.ValueType.String);

    row.cells.A1.value = new Date();
    expect(row.cells.A1.type).to.equal(Enums.ValueType.Date);
    expect(row.cells.A1.effectiveType).to.equal(Enums.ValueType.Date);

    row.cells.A1.value = {formula: 'A2', result: 5};
    expect(row.cells.A1.type).to.deep.equal(Enums.ValueType.Formula);
    expect(row.cells.A1.effectiveType).to.equal(Enums.ValueType.Number);

    row.cells.A1.value = {formula: 'A2', result: 'Hello, World!'};
    expect(row.cells.A1.type).to.deep.equal(Enums.ValueType.Formula);
    expect(row.cells.A1.effectiveType).to.equal(Enums.ValueType.String);

    row.cells.A1.value = {hyperlink: 'http://www.link.com', text: 'www.link.com'};
    expect(row.cells.A1.type).to.deep.equal(Enums.ValueType.Hyperlink);
    expect(row.cells.A1.effectiveType).to.equal(Enums.ValueType.Hyperlink);
  });
});