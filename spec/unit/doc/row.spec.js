var expect = require('chai').expect;

var Column = require("../../../lib/doc/column");
var Row = require("../../../lib/doc/row");
var Enums = require("../../../lib/doc/enums");

function createSheetMock() {
  return {
    _keys: {},
    _cells: {},
    rows: [],
    columns: [],
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
    getCell: function(address) {
      return this._cells[address];
    }
  };
}

describe("Row", function() {
  it("stores cells", function() {
    var sheet = createSheetMock();
    sheet._keys.name = new Column(sheet, 1);

    var row1 = sheet.getRow(1);
    expect(row1.number).to.equal(1);
    expect(row1.hasValues).to.not.be.ok;

    var a1 = row1.getCell(1);
    expect(a1.address).to.equal("A1");
    expect(a1.type).to.equal(Enums.ValueType.Null);
    expect(row1.hasValues).to.not.be.ok;

    expect(row1.getCell("A")).to.equal(a1);
    expect(row1.getCell("name")).to.equal(a1);

    a1.value = 5;
    expect(a1.type).to.equal(Enums.ValueType.Number);
    expect(row1.hasValues).to.be.ok;

    var b1 = row1.getCell(2);
    expect(b1.address).to.equal("B1");
    expect(b1.type).to.equal(Enums.ValueType.Null);
    expect(a1.type).to.equal(Enums.ValueType.Number);

    b1.value = "Hello, World!";
    var d1 = row1.getCell(4);
    d1.value = { hyperlink:"http://www.hyperlink.com", text: "www.hyperlink.com" };

    var values = [,5,"Hello, World!",,{hyperlink:"http://www.hyperlink.com", text: "www.hyperlink.com"}];
    expect(row1.values).to.deep.equal(values);
    expect(row1.dimensions).to.deep.equal({min:1,max:4});

    var count = 0;
    row1.eachCell(function(cell, colNumber) {
      expect(cell.type).to.not.equal(Enums.ValueType.Null);
      switch(cell.type) {
        case Enums.ValueType.Hyperlink:
          expect(cell.value).to.deep.equal(values[colNumber]);
          break;
        default:
          expect(cell.value).to.equal(values[colNumber]);
          break;
      }
      count++;
    });
    expect(count).to.equal(3); //eachCell should only cover non-null cells

    var row2 = sheet.getRow(2);
    expect(row2.dimensions).to.be.null;
  });

  it("stores values by whole row", function() {
    var sheet = createSheetMock();
    sheet._keys = {
      id: new Column(sheet, 1),
      name: new Column(sheet, 2),
      dob: new Column(sheet, 3)
    };

    var now = new Date();

    var row1 = sheet.getRow(1);

    // set values by contiguous array
    row1.values = [5, "Hello, World!", null];
    expect(row1.getCell(1).value).to.equal(5);
    expect(row1.getCell(2).value).to.equal("Hello, World!");
    expect(row1.getCell(3).value).to.be.null;
    expect(row1.values).to.deep.equal([, 5, "Hello, World!"]);

    // set values by sparse array
    var values = [];
    values[1] = 7;
    values[3] = "Not Null!";
    values[5] = now;
    row1.values = values;
    expect(row1.getCell(1).value).to.equal(7);
    expect(row1.getCell(2).value).to.be.null;
    expect(row1.getCell(3).value).to.equal("Not Null!");
    expect(row1.getCell(5).type).to.equal(Enums.ValueType.Date);
    expect(row1.values).to.deep.equal([, 7, , "Not Null!", , now]);

    // set values by object
    row1.values = {
      id: 9,
      name: "Dobbie",
      dob: now
    };
    expect(row1.getCell(1).value).to.equal(9);
    expect(row1.getCell(2).value).to.equal("Dobbie");
    expect(row1.getCell(3).type).to.equal(Enums.ValueType.Date);
    expect(row1.getCell(5).value).to.be.null;
    expect(row1.values).to.deep.equal([, 9, "Dobbie", now]);
  });

  it("iterates over cells", function() {
    var sheet = createSheetMock();
    var row1 = sheet.getRow(1);

    row1.getCell(1).value = 1;
    row1.getCell(2).value = 2;
    row1.getCell(4).value = 4;
    row1.getCell(6).value = 6;
    row1.eachCell(function(cell, colNumber) {
      expect(colNumber).to.not.equal(3);
      expect(colNumber).to.not.equal(5);
      expect(cell.value).to.equal(colNumber);
    });

    var count = 1;
    row1.eachCell({includeEmpty: true}, function(cell, colNumber) {
      expect(colNumber).to.equal(count++);
    });
    expect(count).to.equal(7);
  });

  it("builds a model", function() {
    var sheet = createSheetMock();
    var row1 = sheet.getRow(1);
    row1.getCell(1).value = 5;
    row1.getCell(2).value = "Hello, World!";
    row1.getCell(4).value = { hyperlink:"http://www.hyperlink.com", text: "www.hyperlink.com" };
    row1.getCell(5).value = null;
    row1.height = 50;

    expect(row1.model).to.deep.equal({
      cells:[
        {address:"A1",type:Enums.ValueType.Number,value:5,style:{}},
        {address:"B1",type:Enums.ValueType.String,value:"Hello, World!",style:{}},
        {address:"D1",type:Enums.ValueType.Hyperlink,text:"www.hyperlink.com",hyperlink:"http://www.hyperlink.com",style:{}}
      ],
      number: 1,
      min: 1,
      max: 4,
      height: 50,
      hidden: false,
      style: {}
    });

    var row2 = sheet.getRow(2);
    expect(row2.model).to.be.null;
  });

  it("builds from model", function() {
    var sheet = createSheetMock();
    var row1 = sheet.getRow(1);
    row1.model = {
      cells:[
        {address:"A1",type:Enums.ValueType.Number,value:5},
        {address:"B1",type:Enums.ValueType.String,value:"Hello, World!"},
        {address:"D1",type:Enums.ValueType.Hyperlink,text:"www.hyperlink.com",hyperlink:"http://www.hyperlink.com"}
      ],
      number: 1,
      min: 1,
      max: 4,
      height: 32.5
    };

    expect(row1.dimensions).to.deep.equal({min:1,max:4});
    expect(row1.values).to.deep.equal([,5,"Hello, World!", , {hyperlink:"http://www.hyperlink.com", text: "www.hyperlink.com"}]);
    expect(row1.getCell(1).type).to.equal(Enums.ValueType.Number);
    expect(row1.getCell(1).value).to.equal(5);
    expect(row1.getCell(2).type).to.equal(Enums.ValueType.String);
    expect(row1.getCell(2).value).to.equal("Hello, World!");
    expect(row1.getCell(4).type).to.equal(Enums.ValueType.Hyperlink);
    expect(row1.getCell(4).value).to.deep.equal({hyperlink:"http://www.hyperlink.com", text: "www.hyperlink.com"});
    expect(row1.getCell(5).type).to.equal(Enums.ValueType.Null);
    expect(row1.height - 32.5).to.be.below(0.00000001);
  });
});