var expect = require('chai').expect;

var Column = require('../../../lib/doc/column');
var createSheetMock = require('../../utils/index').createSheetMock;

describe('Column', function() {
  it('creates by defn', function() {
    var sheet = createSheetMock();

    sheet.addColumn(1, {
      header: 'Col 1',
      key: 'id1',
      width: 10
    });

    expect(sheet.getColumn(1).header).to.equal('Col 1');
    expect(sheet.getColumn(1).headers).to.deep.equal(['Col 1']);
    expect(sheet.getCell(1, 1).value).to.equal('Col 1');
    expect(sheet.getColumn('id1')).to.equal(sheet.getColumn(1));
    
    sheet.getRow(2).values = { id1: 'Hello, World!' };
    expect(sheet.getCell(2, 1).value).to.equal('Hello, World!');
  });

  it('maintains properties', function() {
    var sheet = createSheetMock();

    var column = sheet.addColumn(1);

    column.key = 'id1';
    expect(sheet._keys.id1).to.equal(column);

    expect(column.number).to.equal(1);
    expect(column.letter).to.equal('A');

    column.header = 'Col 1';
    expect(sheet.getColumn(1).header).to.equal('Col 1');
    expect(sheet.getColumn(1).headers).to.deep.equal(['Col 1']);
    expect(sheet.getCell(1, 1).value).to.equal('Col 1');

    column.header = ['Col A1', 'Col A2'];
    expect(sheet.getColumn(1).header).to.deep.equal(['Col A1', 'Col A2']);
    expect(sheet.getColumn(1).headers).to.deep.equal(['Col A1', 'Col A2']);
    expect(sheet.getCell(1, 1).value).to.equal('Col A1');
    expect(sheet.getCell(2, 1).value).to.equal('Col A2');

    sheet.getRow(3).values = { id1: 'Hello, World!' };
    expect(sheet.getCell(3, 1).value).to.equal('Hello, World!');
  });

  it('creates model', function() {
    var sheet = createSheetMock();

    sheet.addColumn(1, {
      header: 'Col 1',
      key: 'id1',
      width: 10
    });
    sheet.addColumn(2, {
      header: 'Col 2',
      key: 'name',
      width: 10
    });
    sheet.addColumn(3, {
      header: 'Col 2',
      key: 'dob',
      width: 10,
      outlineLevel: 1
    });

    var model = Column.toModel(sheet.columns);
    expect(model.length).to.equal(2);

    expect(model[0].width).to.equal(10);
    expect(model[0].outlineLevel).to.equal(0);
    expect(model[0].collapsed).to.equal(false);

    expect(model[1].width).to.equal(10);
    expect(model[1].outlineLevel).to.equal(1);
    expect(model[1].collapsed).to.equal(true);
  });
});
