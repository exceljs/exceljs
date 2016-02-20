var expect = require("chai").expect;

var Column = require("../../lib/column");
var Row = require("../../lib/row");

function createSheetMock() {
  return {
    _keys: {},
    _cells: {},
    columns: [],
    rows: [],

    getColumn: function(colNumber) {
      var column = this.columns[colNumber-1];
      if (!column) {
        column = this.columns[colNumber-1] = new Column(this, colNumber);
      }
      return column;
    },
    getRow: function(rowNumber) {
      var row = this.rows[rowNumber-1];
      if (!row) {
        row = this.rows[rowNumber-1] = new Row(this, rowNumber);
      }
      return row;
    },
    getCell: function(rowNumber, colNumber) {
      return this.getRow(rowNumber).getCell(colNumber);
    }
  };
}

describe("Column", function() {
  it("creates by defn", function() {
    var sheet = createSheetMock();

    sheet.columns[0] = new Column(sheet, 1, {
      header: "Col 1",
      key: "id1",
      width: 10
    });

    expect(sheet.getColumn(1).header).to.equal("Col 1");
    expect(sheet.getColumn(1).headers).to.deep.equal(["Col 1"]);
    expect(sheet.getCell(1,1).value).to.equal("Col 1");

    sheet.getRow(2).values = { id1: "Hello, World!" };
    expect(sheet.getCell(2,1).value).to.equal("Hello, World!");
  });

  it("maintains properties", function() {
    var sheet = createSheetMock();

    var column = sheet.columns[0] = new Column(sheet, 1);

    column.key = "id1";
    expect(sheet._keys["id1"]).to.equal(column);

    column.header = "Col 1";
    expect(sheet.getColumn(1).header).to.equal("Col 1");
    expect(sheet.getColumn(1).headers).to.deep.equal(["Col 1"]);
    expect(sheet.getCell(1,1).value).to.equal("Col 1");

    column.header = ["Col A1","Col A2"];
    expect(sheet.getColumn(1).header).to.deep.equal(["Col A1","Col A2"]);
    expect(sheet.getColumn(1).headers).to.deep.equal(["Col A1","Col A2"]);
    expect(sheet.getCell(1,1).value).to.equal("Col A1");
    expect(sheet.getCell(2,1).value).to.equal("Col A2");

    sheet.getRow(3).values = { id1: "Hello, World!" };
    expect(sheet.getCell(3,1).value).to.equal("Hello, World!");
  });

});