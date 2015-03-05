var fs = require("fs");
var _ = require("underscore");
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
        numFmt2: "[Green]#,##0 ;[Red](#,##0)",
        numFmtDate: "dd, mmm yyyy"
    };
    var fonts = {
        arialBlackUI14: { name: "Arial Black", family: 2, size: 14, underline: true, italic: true },
        comicSansUdB16: { name: "Comic Sans MS", family: 4, size: 16, underline: "double", bold: true },
        broadwayRedOutline20: { name: "Broadway", family: 5, size: 20, outline: true, color: { argb:"FFFF0000"}}
    };
    var alignments = [
        { text: "Top Left", alignment: { horizontal: "left", vertical: "top" } },
        { text: "Middle Centre", alignment: { horizontal: "center", vertical: "middle" } },
        { text: "Bottom Right", alignment: { horizontal: "right", vertical: "bottom" } },
        { text: "Wrap Text", alignment: { wrapText: true } },
        { text: "Indent 1", alignment: { indent: 1 } },
        { text: "Indent 2", alignment: { indent: 2 } },
        { text: "Rotate 15", alignment: { horizontal: "right", vertical: "bottom", textRotation: 15 } },
        { text: "Rotate 30", alignment: { horizontal: "right", vertical: "bottom", textRotation: 30 } },
        { text: "Rotate 45", alignment: { horizontal: "right", vertical: "bottom", textRotation: 45 } },
        { text: "Rotate 60", alignment: { horizontal: "right", vertical: "bottom", textRotation: 60 } },
        { text: "Rotate 75", alignment: { horizontal: "right", vertical: "bottom", textRotation: 75 } },
        { text: "Rotate 90", alignment: { horizontal: "right", vertical: "bottom", textRotation: 90 } },
        { text: "Rotate -15", alignment: { horizontal: "right", vertical: "bottom", textRotation: -55 } },
        { text: "Rotate -30", alignment: { horizontal: "right", vertical: "bottom", textRotation: -30 } },
        { text: "Rotate -45", alignment: { horizontal: "right", vertical: "bottom", textRotation: -45 } },
        { text: "Rotate -60", alignment: { horizontal: "right", vertical: "bottom", textRotation: -60 } },
        { text: "Rotate -75", alignment: { horizontal: "right", vertical: "bottom", textRotation: -75 } },
        { text: "Rotate -90", alignment: { horizontal: "right", vertical: "bottom", textRotation: -90 } },
        { text: "Vertical Text", alignment: { horizontal: "right", vertical: "bottom", textRotation: "vertical" } }
    ];
    var badAlignments = [
        { text: "Rotate -91", alignment: { textRotation: -91 } },
        { text: "Rotate 91", alignment: { textRotation: 91 } },
        { text: "Indent -1", alignment: { indent: -1 } },
        { text: "Blank", alignment: {  } }
    ];

    var borders = {
        thin: { top: {style:"thin"}, left: {style:"thin"}, bottom: {style:"thin"}, right: {style:"thin"}},
        doubleRed: { top: {style:"double", color: {argb:"FFFF0000"}}, left: {style:"double", color: {argb:"FFFF0000"}}, bottom: {style:"double", color: {argb:"FFFF0000"}}, right: {style:"double", color: {argb:"FFFF0000"}}},
        thickRainbow: {
            top: {style:"double", color: {argb:"FFFF00FF"}},
            left: {style:"double", color: {argb:"FF00FFFF"}},
            bottom: {style:"double", color: {argb:"FF00FF00"}},
            right: {style:"double", color: {argb:"FF00FF"}},
            diagonal: {style:"double", color: {argb:"FFFFFF00"}, up: true, down: true},
        }
    };

    var fills = {
        redDarkVertical: {type: "pattern", pattern:"darkVertical", fgColor:{argb:"FFFF0000"}},
        redGreenDarkTrellis: {type: "pattern", pattern:"darkTrellis",
            fgColor:{argb:"FFFF0000"}, bgColor:{argb:"FF00FF00"}},
        blueWhiteHGrad: {type: "gradient", gradient: "angle", degree: 0,
            stops: [{position:0, color:{argb:"FF0000FF"}},{position:1, color:{argb:"FFFFFFFF"}}]},
        rgbPathGrad: {type: "gradient", gradient: "path", center:{left:0.5,top:0.5},
            stops: [
                {position:0, color:{argb:"FFFF0000"}},
                {position:0.5, color:{argb:"FF00FF00"}},
                {position:1, color:{argb:"FF0000FF"}}
            ]
        }
    };

    
    var createTestBook = function(checkBadAlignments) {
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
        ws.getCell("A4").border = borders.thin;
        ws.getCell("C4").value = 1.5;
        ws.getCell("C4").numFmt = testValues.numFmt2;
        ws.getCell("C4").border = borders.doubleRed;
        ws.getCell("E4").value = 1.5;
        ws.getCell("E4").border = borders.thickRainbow;
        
        // test fonts and formats
        ws.getCell("A5").value = testValues.str;
        ws.getCell("A5").font = fonts.arialBlackUI14;
        ws.getCell("B5").value = testValues.str;
        ws.getCell("B5").font = fonts.broadwayRedOutline20;
        ws.getCell("C5").value = testValues.str;
        ws.getCell("C5").font = fonts.comicSansUdB16;
        
        ws.getCell("D5").value = 1.6;
        ws.getCell("D5").numFmt = testValues.numFmt1;
        ws.getCell("D5").font = fonts.arialBlackUI14;
        
        ws.getCell("E5").value = 1.6;
        ws.getCell("E5").numFmt = testValues.numFmt2;
        ws.getCell("E5").font = fonts.broadwayRedOutline20;
        
        ws.getCell("F5").value = testValues.date;
        ws.getCell("F5").numFmt = testValues.numFmtDate;
        ws.getCell("F5").font = fonts.comicSansUdB16;
        
        ws.getRow(6).height = 42;
        _.each(alignments, function(alignment, index) {
            var rowNumber = 6;
            var colNumber = index + 1;
            var cell = ws.getCell(rowNumber, colNumber);
            cell.value = alignment.text;
            cell.alignment = alignment.alignment;
        });
        
        if (checkBadAlignments) {
            _.each(badAlignments, function(alignment, index) {
                var rowNumber = 7;
                var colNumber = index + 1;
                var cell = ws.getCell(rowNumber, colNumber);
                cell.value = alignment.text;
                cell.alignment = alignment.alignment;
            });
        }
        
        var row8 = ws.getRow(8);
        row8.height = 40;
        row8.getCell(1).value = "Blue White Horizontal Gradient";
        row8.getCell(1).fill = fills.blueWhiteHGrad;
        row8.getCell(2).value = "Red Dark Vertical";
        row8.getCell(2).fill = fills.redDarkVertical;
        row8.getCell(3).value = "Red Green Dark Trellis";
        row8.getCell(3).fill = fills.redGreenDarkTrellis;
        row8.getCell(4).value = "RGB Path Gradient";
        row8.getCell(4).fill = fills.rgbPathGrad;
        
        return wb;
    }
    var checkFont = function(cell, font) {
        expect(cell.font).toBeDefined();
        _.each(font, function(item, name) {
            expect(cell.font[name]).toEqual(font[name]);
        });
        _.each(value, function(item, name) {
            expect(font[name]).not.toBeDefined();
        });
    };
    var checkTestBook = function(wb, checkBadAlignments) {
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
        expect(ws.getCell("A4").type).toEqual(Excel.ValueType.Number);
        expect(ws.getCell("A4").border).toEqual(borders.thin);
        expect(ws.getCell("C4").numFmt).toEqual(testValues.numFmt2);
        expect(ws.getCell("C4").type).toEqual(Excel.ValueType.Number);
        expect(ws.getCell("C4").border).toEqual(borders.doubleRed);
        expect(ws.getCell("E4").border).toEqual(borders.thickRainbow);
        
        // test fonts and formats
        expect(ws.getCell("A5").value).toEqual(testValues.str);
        expect(ws.getCell("A5").type).toEqual(Excel.ValueType.String);
        expect(ws.getCell("A5").font).toEqual(fonts.arialBlackUI14);
        expect(ws.getCell("B5").value).toEqual(testValues.str);
        expect(ws.getCell("B5").type).toEqual(Excel.ValueType.String);
        expect(ws.getCell("B5").font).toEqual(fonts.broadwayRedOutline20);
        expect(ws.getCell("C5").value).toEqual(testValues.str);
        expect(ws.getCell("C5").type).toEqual(Excel.ValueType.String);
        expect(ws.getCell("C5").font).toEqual(fonts.comicSansUdB16);
        
        expect(Math.abs(ws.getCell("D5").value - 1.6)).toBeLessThan(0.00000001);
        expect(ws.getCell("D5").type).toEqual(Excel.ValueType.Number);
        expect(ws.getCell("D5").numFmt).toEqual(testValues.numFmt1);
        expect(ws.getCell("D5").font).toEqual(fonts.arialBlackUI14);
        
        expect(Math.abs(ws.getCell("E5").value - 1.6)).toBeLessThan(0.00000001);
        expect(ws.getCell("E5").type).toEqual(Excel.ValueType.Number);
        expect(ws.getCell("E5").numFmt).toEqual(testValues.numFmt2);
        expect(ws.getCell("E5").font).toEqual(fonts.broadwayRedOutline20);
        
        expect(Math.abs(ws.getCell("F5").value.getTime() - testValues.date.getTime())).toBeLessThan(3);
        expect(ws.getCell("F5").type).toEqual(Excel.ValueType.Date);
        expect(ws.getCell("F5").numFmt).toEqual(testValues.numFmtDate);
        expect(ws.getCell("F5").font).toEqual(fonts.comicSansUdB16);
        
        expect(ws.getRow(5).height).not.toBeDefined();
        expect(ws.getRow(6).height).toEqual(42);
        _.each(alignments, function(alignment, index) {
            var rowNumber = 6
            var colNumber = index + 1;
            var cell = ws.getCell(rowNumber, colNumber);
            expect(cell.value).toEqual(alignment.text);
            expect(cell.alignment).toEqual(alignment.alignment);
        });
        
        if (checkBadAlignments) {
            _.each(badAlignments, function(alignment, index) {
                var rowNumber = 7;
                var colNumber = index + 1;
                var cell = ws.getCell(rowNumber, colNumber);
                expect(cell.value).toEqual(alignment.text);
                expect(cell.alignment).not.toBeDefined();
            });
        }
        
        var row8 = ws.getRow(8);
        expect(row8.height).toEqual(40);
        expect(row8.getCell(1).fill).toEqual(fills.blueWhiteHGrad);
        expect(row8.getCell(2).fill).toEqual(fills.redDarkVertical);
        expect(row8.getCell(3).fill).toEqual(fills.redGreenDarkTrellis);
        expect(row8.getCell(4).fill).toEqual(fills.rgbPathGrad);
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
        
        var wb = createTestBook(true);
        //fs.writeFileSync("./testmodel.json", JSON.stringify(wb.model, null, "    "));
        
        wb.xlsx.writeFile("./wb.test.xlsx")
            .then(function() {
                var wb2 = new Excel.Workbook();
                return wb2.xlsx.readFile("./wb.test.xlsx");
            })
            .then(function(wb2) {
                checkTestBook(wb2, true);
            })
            .finally(function() {
                fs.unlink("./wb.test.xlsx", function(error) {
                    expect(error && error.message).toBeFalsy();
                    done();
                });
            });
    });
    
    it("serializes and deserializes to file properly", function(done) {
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
                    var ws = wb.getWorksheet("sheet" + i);
                    expect(ws).toBeDefined();
                    expect(ws.getCell("A1").value).toEqual(i);
                }
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
    
    it("throws an error when file not found", function(done) {
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
    })
});
