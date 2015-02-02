var fs = require("fs");
var Excel = require("../excel");
var utils = require("./testutils");

describe("Workbook", function() {
    var testValues = {
        num: 7,
        str: "Hello, World!",
        str2: '<a href="www.whatever.com">Talk to the H&</a>',
        date: new Date(),
        formulas: [
            {formula: "A1", result: 7},
            {formula: "A2"}
        ],
        hyperlink: {hyperlink: "http://www.link.com", text: "www.link.com"},
        numFmt1: "# ?/?",
        numFmt2: "[Green]#,##0 ;[Red](#,##0)"
    };
    var createTestBook = function() {
        var wb = new Excel.Workbook()
        var ws = wb.addWorksheet("blort");
        
        ws.getCell("A1").value = 7;
        ws.getCell("B1").value = testValues.str;
        ws.getCell("C1").value = testValues.date;
        ws.getCell("D1").value = testValues.formulas[0];
        ws.getCell("E1").value = testValues.formulas[1];
        ws.getCell("F1").value = testValues.hyperlink;
        ws.getCell("G1").value = testValues.str2;

        // merge cell square with numerical value
        ws.getCell("A2").value = 5;
        ws.mergeCells("A2:B3");
        
        // merge cell squalre with null value
        ws.mergeCells("C2:D3");
        
        ws.getCell("A4").value = 1.5;
        ws.getCell("A4").numFmt = testValues.numFmt1;
        ws.getCell("B4").value = 1.5;
        ws.getCell("B4").numFmt = testValues.numFmt2;
        
        return wb;
    }
    var checkTestBook = function(wb) {
        expect(wb).toBeDefined();
        
        var ws = wb.getWorksheet("blort");
        expect(ws).toBeDefined();
        
        expect(ws.getCell("A1").value).toEqual(7);
        expect(ws.getCell("A1").type).toEqual(Excel.ValueType.Number);
        expect(ws.getCell("B1").value).toEqual(testValues.str);
        expect(ws.getCell("B1").type).toEqual(Excel.ValueType.String);
        expect(Math.abs(ws.getCell("C1").value.getTime() - testValues.date.getTime())).toBeLessThan(3);
        expect(ws.getCell("C1").type).toEqual(Excel.ValueType.Date);
        expect(ws.getCell("D1").value).toEqual(testValues.formulas[0]);
        expect(ws.getCell("D1").type).toEqual(Excel.ValueType.Formula);
        expect(ws.getCell("E1").value).toEqual(testValues.formulas[1]);
        expect(ws.getCell("E1").type).toEqual(Excel.ValueType.Formula);
        expect(ws.getCell("F1").value).toEqual(testValues.hyperlink);
        expect(ws.getCell("F1").type).toEqual(Excel.ValueType.Hyperlink);
        expect(ws.getCell("G1").value).toEqual(testValues.str2);
        
        // A2:B3
        expect(ws.getCell("A2").value).toEqual(5);
        expect(ws.getCell("A2").type).toEqual(Excel.ValueType.Number);
        expect(ws.getCell("A2").master).toBe(ws.getCell("A2"));
        
        expect(ws.getCell("A3").value).toEqual(5);
        expect(ws.getCell("A3").type).toEqual(Excel.ValueType.Merge);
        expect(ws.getCell("A3").master).toBe(ws.getCell("A2"));
        
        expect(ws.getCell("B2").value).toEqual(5);
        expect(ws.getCell("B2").type).toEqual(Excel.ValueType.Merge);
        expect(ws.getCell("B2").master).toBe(ws.getCell("A2"));
        
        expect(ws.getCell("B3").value).toEqual(5);
        expect(ws.getCell("B3").type).toEqual(Excel.ValueType.Merge);
        expect(ws.getCell("B3").master).toBe(ws.getCell("A2"));
        
        // C2:D3
        expect(ws.getCell("C2").value).toBeNull();
        expect(ws.getCell("C2").type).toEqual(Excel.ValueType.Null);
        expect(ws.getCell("C2").master).toBe(ws.getCell("C2"));
        
        expect(ws.getCell("D2").value).toBeNull();
        expect(ws.getCell("D2").type).toEqual(Excel.ValueType.Merge);
        expect(ws.getCell("D2").master).toBe(ws.getCell("C2"));
        
        expect(ws.getCell("C3").value).toBeNull();
        expect(ws.getCell("C3").type).toEqual(Excel.ValueType.Merge);
        expect(ws.getCell("C3").master).toBe(ws.getCell("C2"));
        
        expect(ws.getCell("D3").value).toBeNull();
        expect(ws.getCell("D3").type).toEqual(Excel.ValueType.Merge);
        expect(ws.getCell("D3").master).toBe(ws.getCell("C2"));
        
        expect(ws.getCell("A4").numFmt).toEqual(testValues.numFmt1);
        expect(ws.getCell("B4").numFmt).toEqual(testValues.numFmt2);
    }
    
    it("creates sheets with correct names", function() {
        var wb = new Excel.Workbook();
        var ws1 = wb.addWorksheet("Hello, World!");
        expect(ws1.name).toEqual("Hello, World!");
        
        var ws2 = wb.addWorksheet();
        expect(ws2.name).toMatch(/sheet\d+/);
    });
    it("serialises and deserialises by model", function(done) {
        
        var wb = createTestBook();
        
        utils.cloneByModel(wb, Excel.Workbook)
            .then(function(wb2) {
                checkTestBook(wb2);
            })
            .catch(function(error) {
                console.log(error.message);
                expect(error && error.message).toBeUndefined();
            })
            .finally(function() {
                done();
            });
    });
    
    it("serializes and deserializes to file properly", function(done) {
        
        var wb = createTestBook();
        //fs.writeFileSync("./testmodel.json", JSON.stringify(wb.model, null, "    "));
        
        wb.xlsx.writeFile("./wb.test.xlsx")
            .then(function() {
                var wb2 = new Excel.Workbook();
                return wb2.xlsx.readFile("./wb.test.xlsx");
            })
            .then(function(wb2) {
                checkTestBook(wb2);
            })
            .finally(function() {
                fs.unlink("./wb.test.xlsx", function(error) {
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
});

describe("Merge Cells", function() {
    it("references the same top-left value", function() {
        var wb = new Excel.Workbook()
        var ws = wb.addWorksheet("blort");
        
        // initial values
        ws.getCell("A1").value = "A1";
        ws.getCell("B1").value = "B1";
        ws.getCell("A2").value = "A2";
        ws.getCell("B2").value = "B2";
        
        ws.mergeCells("A1:B2");
        
        expect(ws.getCell("A1").value).toEqual("A1");
        expect(ws.getCell("B1").value).toEqual("A1");
        expect(ws.getCell("A2").value).toEqual("A1");
        expect(ws.getCell("B2").value).toEqual("A1");
        
        expect(ws.getCell("A1").type).toEqual(Excel.ValueType.String);
        expect(ws.getCell("B1").type).toEqual(Excel.ValueType.Merge);
        expect(ws.getCell("A2").type).toEqual(Excel.ValueType.Merge);
        expect(ws.getCell("B2").type).toEqual(Excel.ValueType.Merge);
    });

    it("does not allow overlapping merges", function() {
        var wb = new Excel.Workbook()
        var ws = wb.addWorksheet("blort");
        
        ws.mergeCells("B2:C3");
        
        // intersect four corners
        expect(function() { ws.mergeCells("A1:B2"); }).toThrow();
        expect(function() { ws.mergeCells("C1:D2"); }).toThrow();
        expect(function() { ws.mergeCells("C3:D4"); }).toThrow();
        expect(function() { ws.mergeCells("A3:B4"); }).toThrow();
        
        // enclosing
        expect(function() { ws.mergeCells("A1:D4"); }).toThrow();
    });
});
