var fs = require("fs");
var _ = require("underscore");
var Excel = require("../excel");
var testutils = require("./testutils");

// =============================================================================
// Tests

describe("Workbook", function() {

  it("creates sheets and saves with correct names", function(done) {
    var wb = new Excel.Workbook();
    var ws1 = wb.addWorksheet("Hello, World!");
    expect(ws1.name).toEqual("Hello, World!");

    var ws2 = wb.addWorksheet();
    expect(ws2.name).toMatch(/sheet\d+/);

    var ws3 = wb.addWorksheet("This & That");

    wb.xlsx.writeFile("./wb.test.xlsx")
      .then(function() {
        var wb2 = new Excel.Workbook();
        return wb2.xlsx.readFile("./wb.test.xlsx");
      })
      .then(function(wb2) {
        expect(wb2.getWorksheet("Hello, World!")).toBeDefined();
        expect(wb2.getWorksheet("This & That")).toBeDefined();
      })
      .catch(function(error) {
        expect(error && error.message).toBeFalsy();
      })
      .finally(function() {
        fs.unlink("./wb.test.xlsx", function(error) {
          expect(error && error.message).toBeFalsy();
          done();
        });
      });
  });

  it("serialises and deserialises by model", function(done) {
    var wb = testutils.createTestBook(false, Excel.Workbook);

    testutils.cloneByModel(wb, Excel.Workbook)
      .then(function(wb2) {
        testutils.checkTestBook(wb2, "model");
      })
      .catch(function(error) {
        console.log(error.message);
        expect(error && error.message).toBeUndefined();
      })
      .finally(function() {
        done();
      });
  });

  it("serializes and deserializes to xlsx file properly", function(done) {

    var wb = testutils.createTestBook(true, Excel.Workbook);
    //fs.writeFileSync("./testmodel.json", JSON.stringify(wb.model, null, "    "));

    wb.xlsx.writeFile("./wb.test.xlsx")
      .then(function() {
        var wb2 = new Excel.Workbook();
        return wb2.xlsx.readFile("./wb.test.xlsx");
      })
      .then(function(wb2) {
        testutils.checkTestBook(wb2, "xlsx", true);
      })
      .catch(function(error) {
        expect(error && error.message).toBeFalsy();
      })
      .finally(function() {
        fs.unlink("./wb.test.xlsx", function(error) {
          expect(error && error.message).toBeFalsy();
          done();
        });
      });
  });

  it("serializes row styles and columns properly", function(done) {
    var wb = new Excel.Workbook();
    var ws = wb.addWorksheet("blort");

    ws.columns = [
      { header: "A1", width: 10 },
      { header: "B1", width: 20, style: { font: testutils.styles.fonts.comicSansUdB16, alignment: testutils.styles.alignments[1].alignment } },
      { header: "C1", width: 30 }
    ];

    ws.getRow(2).font = testutils.styles.fonts.broadwayRedOutline20;

    ws.getCell("A2").value = "A2";
    ws.getCell("B2").value = "B2";
    ws.getCell("C2").value = "C2";
    ws.getCell("A3").value = "A3";
    ws.getCell("B3").value = "B3";
    ws.getCell("C3").value = "C3";

    wb.xlsx.writeFile("./wb.test.xlsx")
      .then(function() {
        var wb2 = new Excel.Workbook();
        return wb2.xlsx.readFile("./wb.test.xlsx");
      })
      .then(function(wb2) {
        var ws2 = wb2.getWorksheet("blort");
        _.each(["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"], function(address) {
          expect(ws2.getCell(address).value).toEqual(address);
        });
        expect(ws2.getCell("B1").font).toEqual(testutils.styles.fonts.comicSansUdB16);
        expect(ws2.getCell("B1").alignment).toEqual(testutils.styles.alignments[1].alignment);
        expect(ws2.getCell("A2").font).toEqual(testutils.styles.fonts.broadwayRedOutline20);
        expect(ws2.getCell("B2").font).toEqual(testutils.styles.fonts.broadwayRedOutline20);
        expect(ws2.getCell("C2").font).toEqual(testutils.styles.fonts.broadwayRedOutline20);
        expect(ws2.getCell("B3").font).toEqual(testutils.styles.fonts.comicSansUdB16);
        expect(ws2.getCell("B3").alignment).toEqual(testutils.styles.alignments[1].alignment);

        expect(ws2.getColumn(2).font).toEqual(testutils.styles.fonts.comicSansUdB16);
        expect(ws2.getColumn(2).alignment).toEqual(testutils.styles.alignments[1].alignment);

        expect(ws2.getRow(2).font).toEqual(testutils.styles.fonts.broadwayRedOutline20);
      })
      .catch(function(error) {
        expect(error && error.message).toBeFalsy();
      })
      .finally(function() {
        fs.unlink("./wb.test.xlsx", function(error) {
          expect(error && error.message).toBeFalsy();
          done();
        });
      });
  });

  it("serializes and deserializes a lot of sheets to xlsx file properly", function(done) {
    var wb = new Excel.Workbook();
    var numSheets = 90;
    // add numSheets sheets
    for (i = 1; i <= numSheets; i++) {
      var ws = wb.addWorksheet("sheet" + i);
      ws.getCell("A1").value = i;
    }
    wb.xlsx.writeFile("./wb.test.xlsx")
      .then(function() {
        var wb2 = new Excel.Workbook();
        return wb2.xlsx.readFile("./wb.test.xlsx");
      })
      .then(function(wb2) {
        for (i = 1; i <= numSheets; i++) {
          var ws2 = wb2.getWorksheet("sheet" + i);
          expect(ws2).toBeDefined();
          expect(ws2.getCell("A1").value).toEqual(i);
        }
      })
      .catch(function(error) {
        expect(error && error.message).toBeFalsy();
      })
      .finally(function() {
        fs.unlink("./wb.test.xlsx", function(error) {
          expect(error && error.message).toBeFalsy();
          done();
        });
      });
  });

  it("serializes and deserializes to csv file properly", function(done) {

    var wb = testutils.createTestBook(true, Excel.Workbook);
    //fs.writeFileSync("./testmodel.json", JSON.stringify(wb.model, null, "    "));

    wb.csv.writeFile("./wb.test.csv")
      .then(function() {
        var wb2 = new Excel.Workbook();
        return wb2.csv.readFile("./wb.test.csv")
          .then(function() {
            return wb2;
          });
      })
      .then(function(wb2) {
        testutils.checkTestBook(wb2, "csv");
      })
      .catch(function(error) {
        expect(error && error.message).toBeFalsy();
      })
      .finally(function() {
        fs.unlink("./wb.test.csv", function(error) {
          expect(error && error.message).toBeFalsy();
          done();
        });
      });
  });

  it("stores shared string values properly", function() {
    var wb = new Excel.Workbook()
    var ws = wb.addWorksheet("blort");

    ws.getCell("A1").value = "Hello, World!";

    ws.getCell("A2").value = "Hello";
    ws.getCell("B2").value = "World";
    ws.getCell("C2").value = {formula: 'CONCATENATE(A2, ", ", B2, "!")', result: "Hello, World!"};

    ws.getCell("A3").value = ["Hello", "World"].join(", ") + "!";

    // A1 and A3 should reference the same string object
    expect(ws.getCell("A1").value).toBe(ws.getCell("A3").value);

    // A1 and C2 should not reference the same object
    expect(ws.getCell("A1").value).toBe(ws.getCell("C2").value.result);
  });

  it("assigns cell types properly", function() {
    var wb = new Excel.Workbook()
    var ws = wb.addWorksheet("blort");

    // plain number
    ws.getCell("A1").value = 7;

    // simple string
    ws.getCell("B1").value = "Hello, World!";

    // floating point
    ws.getCell("C1").value = 3.14;

    // date-time
    ws.getCell("D1").value = new Date();

    // hyperlink
    ws.getCell("E1").value = {text: "www.google.com", hyperlink:"http://www.google.com"};

    // number formula
    ws.getCell("A2").value = {formula: "A1", result: 7};

    // string formula
    ws.getCell("B2").value = {formula: 'CONCATENATE("Hello", ", ", "World!")', result: "Hello, World!"};

    // date formula
    ws.getCell("C2").value = {formula: "D1", result: new Date()};

    expect(ws.getCell("A1").type).toEqual(Excel.ValueType.Number);
    expect(ws.getCell("B1").type).toEqual(Excel.ValueType.String);
    expect(ws.getCell("C1").type).toEqual(Excel.ValueType.Number);
    expect(ws.getCell("D1").type).toEqual(Excel.ValueType.Date);
    expect(ws.getCell("E1").type).toEqual(Excel.ValueType.Hyperlink);

    expect(ws.getCell("A2").type).toEqual(Excel.ValueType.Formula);
    expect(ws.getCell("B2").type).toEqual(Excel.ValueType.Formula);
    expect(ws.getCell("C2").type).toEqual(Excel.ValueType.Formula);
  });

  it("throws an error when xlsx file not found", function(done) {
    var wb = new Excel.Workbook();
    var success = 0;
    wb.xlsx.readFile("./wb.doesnotexist.xlsx")
      .then(function(wb) {
        success = 1;
      })
      .catch(function(error) {
        success = 2;
        // expect the right kind of error
      })
      .finally(function() {
        expect(success).toEqual(2);
        done();
      });
  });

  it("throws an error when csv file not found", function(done) {
    var wb = new Excel.Workbook();
    var success = 0;
    wb.csv.readFile("./wb.doesnotexist.csv")
      .then(function(wb) {
        success = 1;
      })
      .catch(function(error) {
        success = 2;
        // expect the right kind of error
      })
      .finally(function() {
        expect(success).toEqual(2);
        done();
      });
  });

});

describe("Merge Cells", function() {
  it("serialises and deserialises properly", function() {
    var wb = new Excel.Workbook()
    var ws = wb.addWorksheet("blort");

    // initial values
    ws.getCell("B2").value = "B2";

    ws.mergeCells("B2:C3");

    wb.xlsx.writeFile("./wb.test.xlsx")
      .then(function() {
        var wb2 = new Excel.Workbook();
        return wb2.xlsx.readFile("./wb.test.xlsx");
      })
      .then(function(wb2) {
        var ws2 = wb2.getWorksheet("blort");

        expect(ws2.getCell("B2").value).toEqual("B2");
        expect(ws2.getCell("B3").value).toEqual("B2");
        expect(ws2.getCell("C2").value).toEqual("B2");
        expect(ws2.getCell("C3").value).toEqual("B2");

        expect(ws2.getCell("B2").type).toEqual(Excel.ValueType.String);
        expect(ws2.getCell("B3").type).toEqual(Excel.ValueType.Merge);
        expect(ws2.getCell("C2").type).toEqual(Excel.ValueType.Merge);
        expect(ws2.getCell("C3").type).toEqual(Excel.ValueType.Merge);
      })
      .catch(function(error) {
        expect(error && error.message).toBeFalsy();
      })
      .finally(function() {
        fs.unlink("./wb.test.xlsx", function(error) {
          expect(error && error.message).toBeFalsy();
          done();
        });
      });
  });

  it("serialises and deserialises styles", function() {
    var wb = new Excel.Workbook()
    var ws = wb.addWorksheet("blort");

    // initial values
    var B2 = ws.getCell("B2");
    B2.value = 5;
    B2.style.font = testutils.styles.fonts.broadwayRedOutline20;
    B2.style.border = testutils.styles.borders.doubleRed;
    B2.style.fill = testutils.styles.fills.blueWhiteHGrad;
    B2.style.alignment = testutils.styles.namedAlignments.middleCentre;
    B2.style.numFmt = testutils.styles.numFmts.numFmt1;

    // expecting styles to be copied (see worksheet spec)
    ws.mergeCells("B2:C3");

    wb.xlsx.writeFile("./wb.test.xlsx")
      .then(function() {
        var wb2 = new Excel.Workbook();
        return wb2.xlsx.readFile("./wb.test.xlsx");
      })
      .then(function(wb2) {
        var ws2 = wb2.getWorksheet("blort");

        expect(ws2.getCell("B2").font).toEqual(testutils.styles.fonts.broadwayRedOutline20);
        expect(ws2.getCell("B2").border).toEqual(testutils.styles.borders.doubleRed);
        expect(ws2.getCell("B2").fill).toEqual(testutils.styles.fills.blueWhiteHGrad);
        expect(ws2.getCell("B2").alignment).toEqual(testutils.styles.namedAlignments.middleCentre);
        expect(ws2.getCell("B2").numFmt).toEqual(testutils.styles.numFmts.numFmt1);

        expect(ws2.getCell("B3").font).toEqual(testutils.styles.fonts.broadwayRedOutline20);
        expect(ws2.getCell("B3").border).toEqual(testutils.styles.borders.doubleRed);
        expect(ws2.getCell("B3").fill).toEqual(testutils.styles.fills.blueWhiteHGrad);
        expect(ws2.getCell("B3").alignment).toEqual(testutils.styles.namedAlignments.middleCentre);
        expect(ws2.getCell("B3").numFmt).toEqual(testutils.styles.numFmts.numFmt1);

        expect(ws2.getCell("C2").font).toEqual(testutils.styles.fonts.broadwayRedOutline20);
        expect(ws2.getCell("C2").border).toEqual(testutils.styles.borders.doubleRed);
        expect(ws2.getCell("C2").fill).toEqual(testutils.styles.fills.blueWhiteHGrad);
        expect(ws2.getCell("C2").alignment).toEqual(testutils.styles.namedAlignments.middleCentre);
        expect(ws2.getCell("C2").numFmt).toEqual(testutils.styles.numFmts.numFmt1);

        expect(ws2.getCell("C3").font).toEqual(testutils.styles.fonts.broadwayRedOutline20);
        expect(ws2.getCell("C3").border).toEqual(testutils.styles.borders.doubleRed);
        expect(ws2.getCell("C3").fill).toEqual(testutils.styles.fills.blueWhiteHGrad);
        expect(ws2.getCell("C3").alignment).toEqual(testutils.styles.namedAlignments.middleCentre);
        expect(ws2.getCell("C3").numFmt).toEqual(testutils.styles.numFmts.numFmt1);

      })
      .catch(function(error) {
        expect(error && error.message).toBeFalsy();
      })
      .finally(function() {
        fs.unlink("./wb.test.xlsx", function(error) {
          expect(error && error.message).toBeFalsy();
          done();
        });
      });
  });
});
