var Column = require("../lib/column");
var Row = require("../lib/row");
var Enums = require("../lib/enums");

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
    expect(row1.number).toEqual(1);
    expect(row1.hasValues).not.toBeTruthy();

    var a1 = row1.getCell(1);
    expect(a1.address).toEqual("A1");
    expect(a1.type).toEqual(Enums.ValueType.Null);
    expect(row1.hasValues).not.toBeTruthy();

    expect(row1.getCell("A")).toBe(a1);
    expect(row1.getCell("name")).toBe(a1);

    a1.value = 5;
    expect(a1.type).toEqual(Enums.ValueType.Number);
    expect(row1.hasValues).toBeTruthy();

    var b1 = row1.getCell(2);
    expect(b1.address).toEqual("B1");
    expect(b1.type).toEqual(Enums.ValueType.Null);
    expect(a1.type).toEqual(Enums.ValueType.Number);

    b1.value = "Hello, World!";
    var d1 = row1.getCell(4);
    d1.value = { hyperlink:"http://www.hyperlink.com", text: "www.hyperlink.com" };

    var values = [
      undefined,
      5,
      "Hello, World!",
      undefined,
      {hyperlink:"http://www.hyperlink.com", text: "www.hyperlink.com"}
    ];
    expect(row1.values).toEqual(values);
    expect(row1.dimensions).toEqual({min:1,max:4});

    var count = 0;
    row1.eachCell(function(cell, colNumber) {
      expect(cell.type).not.toEqual(Enums.ValueType.Null);
      expect(cell.value).toEqual(values[colNumber]);
      count++;
    });
    expect(count).toEqual(3); //eachCell should only cover non-null cells

    var row2 = sheet.getRow(2);
    expect(row2.dimensions).toBeNull();
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
    expect(row1.getCell(1).value).toEqual(5);
    expect(row1.getCell(2).value).toEqual("Hello, World!");
    expect(row1.getCell(3).value).toBeNull();
    expect(row1.values).toEqual([undefined, 5, "Hello, World!"]);

    // set values by sparse array
    var values = [];
    values[1] = 7;
    values[3] = "Not Null!";
    values[5] = now;
    row1.values = values;
    expect(row1.getCell(1).value).toEqual(7);
    expect(row1.getCell(2).value).toBeNull();
    expect(row1.getCell(3).value).toEqual("Not Null!");
    expect(row1.getCell(5).type).toEqual(Enums.ValueType.Date);
    expect(row1.values).toEqual([undefined, 7, undefined, "Not Null!", undefined, now]);

    // set values by object
    row1.values = {
      id: 9,
      name: "Dobbie",
      dob: now
    };
    expect(row1.getCell(1).value).toEqual(9);
    expect(row1.getCell(2).value).toEqual("Dobbie");
    expect(row1.getCell(3).type).toEqual(Enums.ValueType.Date);
    expect(row1.getCell(5).value).toBeNull();
    expect(row1.values).toEqual([undefined, 9, "Dobbie", now]);
  });

  it("iterates over cells", function() {
    var sheet = createSheetMock();
    var row1 = sheet.getRow(1);

    row1.getCell(1).value = 1;
    row1.getCell(2).value = 2;
    row1.getCell(4).value = 4;
    row1.getCell(6).value = 6;
    row1.eachCell(function(cell, colNumber) {
      expect(colNumber).not.toEqual(3);
      expect(colNumber).not.toEqual(5);
      expect(cell.value).toEqual(colNumber);
    });

    var count = 1;
    row1.eachCell({includeEmpty: true}, function(cell, colNumber) {
      expect(colNumber).toEqual(count++);
    });
    expect(count).toEqual(7);
  });

  it("builds a model", function() {
    var sheet = createSheetMock();
    var row1 = sheet.getRow(1);
    row1.getCell(1).value = 5;
    row1.getCell(2).value = "Hello, World!";
    row1.getCell(4).value = { hyperlink:"http://www.hyperlink.com", text: "www.hyperlink.com" };
    row1.getCell(5).value = null;
    row1.height = 50;

    expect(row1.model).toEqual({
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
    expect(row2.model).toBeNull();
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

    expect(row1.dimensions).toEqual({min:1,max:4});
    expect(row1.values).toEqual([undefined,5,"Hello, World!", undefined, {hyperlink:"http://www.hyperlink.com", text: "www.hyperlink.com"}]);
    expect(row1.getCell(1).type).toEqual(Enums.ValueType.Number);
    expect(row1.getCell(1).value).toEqual(5);
    expect(row1.getCell(2).type).toEqual(Enums.ValueType.String);
    expect(row1.getCell(2).value).toEqual("Hello, World!");
    expect(row1.getCell(4).type).toEqual(Enums.ValueType.Hyperlink);
    expect(row1.getCell(4).value).toEqual({hyperlink:"http://www.hyperlink.com", text: "www.hyperlink.com"});
    expect(row1.getCell(5).type).toEqual(Enums.ValueType.Null);
    expect(row1.height - 32.5).toBeLessThan(0.00000001);
  });
});