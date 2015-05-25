var Promise = require("bluebird");
var _ = require("underscore");
var MemoryStream = require("memorystream");

var Excel = require("../excel");

var utils = module.exports = {
    cloneByModel: function(thing1, type) {
        var model = thing1.model;
        //console.log(JSON.stringify(model, null, "    "))
        var thing2 = new type();
        thing2.model = model;
        return Promise.resolve(thing2);
    },
    cloneByStream: function(thing1, type, end) {
        var deferred = Promise.defer();
        end = end || "end";
        
        var thing2 = new type();
        var stream = thing2.createInputStream();
        stream.on(end, function() {
            deferred.resolve(thing2);            
        });
        stream.on("error", function(error) {
            deferred.reject(error);
        });
        
        var memStream = new MemoryStream();
        memStream.on("error", function(error) {
            deferred.reject(error);
        });
        memStream.pipe(stream);
        thing1.write(memStream)
            .then(function() {
                memStream.end();
            });
            
        return deferred.promise;
    },
    testValues: {
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
    },
    styles: {
        numFmts: {
            numFmt1: "# ?/?",
            numFmt2: "[Green]#,##0 ;[Red](#,##0)"
        },
        fonts: {
            arialBlackUI14: { name: "Arial Black", family: 2, size: 14, underline: true, italic: true },
            comicSansUdB16: { name: "Comic Sans MS", family: 4, size: 16, underline: "double", bold: true },
            broadwayRedOutline20: { name: "Broadway", family: 5, size: 20, outline: true, color: { argb:"FFFF0000"}}
        },
        alignments: [
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
        ],
        namedAlignments: {
            topLeft: { horizontal: "left", vertical: "top" },
            middleCentre: { horizontal: "center", vertical: "middle" },
            bottomRight: { horizontal: "right", vertical: "bottom" }
        },
        badAlignments: [
            { text: "Rotate -91", alignment: { textRotation: -91 } },
            { text: "Rotate 91", alignment: { textRotation: 91 } },
            { text: "Indent -1", alignment: { indent: -1 } },
            { text: "Blank", alignment: {  } }
        ],
        borders: {
            thin: { top: {style:"thin"}, left: {style:"thin"}, bottom: {style:"thin"}, right: {style:"thin"}},
            doubleRed: { top: {style:"double", color: {argb:"FFFF0000"}}, left: {style:"double", color: {argb:"FFFF0000"}}, bottom: {style:"double", color: {argb:"FFFF0000"}}, right: {style:"double", color: {argb:"FFFF0000"}}},
            thickRainbow: {
                top: {style:"double", color: {argb:"FFFF00FF"}},
                left: {style:"double", color: {argb:"FF00FFFF"}},
                bottom: {style:"double", color: {argb:"FF00FF00"}},
                right: {style:"double", color: {argb:"FF00FF"}},
                diagonal: {style:"double", color: {argb:"FFFFFF00"}, up: true, down: true},
            }
        },
        fills: {
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
        }
    },
    
    createTestBook: function(checkBadAlignments, workbookClass, options) {
        var wb = new workbookClass(options)
        var ws = wb.addWorksheet("blort");
        
        ws.getCell("A1").value = 7;
        ws.getCell("B1").value = utils.testValues.str;
        ws.getCell("C1").value = utils.testValues.date;
        ws.getCell("D1").value = utils.testValues.formulas[0];
        ws.getCell("E1").value = utils.testValues.formulas[1];
        ws.getCell("F1").value = utils.testValues.hyperlink;
        ws.getCell("G1").value = utils.testValues.str2;
        ws.getRow(1).commit();
        
        // merge cell square with numerical value
        ws.getCell("A2").value = 5;
        ws.mergeCells("A2:B3");
        
        // merge cell square with null value
        ws.mergeCells("C2:D3");
        ws.getRow(3).commit();
        
        ws.getCell("A4").value = 1.5;
        ws.getCell("A4").numFmt = utils.testValues.numFmt1;
        ws.getCell("A4").border = utils.styles.borders.thin;
        ws.getCell("C4").value = 1.5;
        ws.getCell("C4").numFmt = utils.testValues.numFmt2;
        ws.getCell("C4").border = utils.styles.borders.doubleRed;
        ws.getCell("E4").value = 1.5;
        ws.getCell("E4").border = utils.styles.borders.thickRainbow;
        ws.getRow(4).commit();
        
        // test fonts and formats
        ws.getCell("A5").value = utils.testValues.str;
        ws.getCell("A5").font = utils.styles.fonts.arialBlackUI14;
        ws.getCell("B5").value = utils.testValues.str;
        ws.getCell("B5").font = utils.styles.fonts.broadwayRedOutline20;
        ws.getCell("C5").value = utils.testValues.str;
        ws.getCell("C5").font = utils.styles.fonts.comicSansUdB16;
        
        ws.getCell("D5").value = 1.6;
        ws.getCell("D5").numFmt = utils.testValues.numFmt1;
        ws.getCell("D5").font = utils.styles.fonts.arialBlackUI14;
        
        ws.getCell("E5").value = 1.6;
        ws.getCell("E5").numFmt = utils.testValues.numFmt2;
        ws.getCell("E5").font = utils.styles.fonts.broadwayRedOutline20;
        
        ws.getCell("F5").value = utils.testValues.date;
        ws.getCell("F5").numFmt = utils.testValues.numFmtDate;
        ws.getCell("F5").font = utils.styles.fonts.comicSansUdB16;
        ws.getRow(5).commit();
        
        ws.getRow(6).height = 42;
        _.each(utils.styles.alignments, function(alignment, index) {
            var rowNumber = 6;
            var colNumber = index + 1;
            var cell = ws.getCell(rowNumber, colNumber);
            cell.value = alignment.text;
            cell.alignment = alignment.alignment;
        });
        ws.getRow(6).commit();
        
        if (checkBadAlignments) {
            _.each(utils.styles.badAlignments, function(alignment, index) {
                var rowNumber = 7;
                var colNumber = index + 1;
                var cell = ws.getCell(rowNumber, colNumber);
                cell.value = alignment.text;
                cell.alignment = alignment.alignment;
            });
        }
        ws.getRow(7).commit();
        
        var row8 = ws.getRow(8);
        row8.height = 40;
        row8.getCell(1).value = "Blue White Horizontal Gradient";
        row8.getCell(1).fill = utils.styles.fills.blueWhiteHGrad;
        row8.getCell(2).value = "Red Dark Vertical";
        row8.getCell(2).fill = utils.styles.fills.redDarkVertical;
        row8.getCell(3).value = "Red Green Dark Trellis";
        row8.getCell(3).fill = utils.styles.fills.redGreenDarkTrellis;
        row8.getCell(4).value = "RGB Path Gradient";
        row8.getCell(4).fill = utils.styles.fills.rgbPathGrad;
        row8.commit();
        
        return wb;
    },
    checkFont: function(cell, font) {
        expect(cell.font).toBeDefined();
        _.each(font, function(item, name) {
            expect(cell.font[name]).toEqual(font[name]);
        });
        _.each(value, function(item, name) {
            expect(font[name]).not.toBeDefined();
        });
    },
    checkTestBook: function(wb, docType, useStyles) {
        var sheetName;
        var checkFormulas, checkMerges, checkStyles, checkBadAlignments;
        var dateAccuracy;
        switch(docType) {
            case "xlsx":
                sheetName = "blort";
                checkFormulas = true;
                checkMerges = true;
                checkStyles = useStyles;
                checkBadAlignments = useStyles;
                dateAccuracy = 3;
                break;
            case "model":
                sheetName = "blort";
                checkFormulas = true;
                checkMerges = true;
                checkStyles = true;
                checkBadAlignments = false;
                dateAccuracy = 3;
                break;
            case "csv":
                sheetName = "sheet1";
                checkFormulas = false;
                checkMerges = false;
                checkStyles = false;
                checkBadAlignments = false;
                dateAccuracy = 1000;
                break;
        }
        
        expect(wb).toBeDefined();
        
        var ws = wb.getWorksheet(sheetName);
        expect(ws).toBeDefined();
        
        expect(ws.getCell("A1").value).toEqual(7);
        expect(ws.getCell("A1").type).toEqual(Excel.ValueType.Number);
        expect(ws.getCell("B1").value).toEqual(utils.testValues.str);
        expect(ws.getCell("B1").type).toEqual(Excel.ValueType.String);
        expect(Math.abs(ws.getCell("C1").value.getTime() - utils.testValues.date.getTime())).toBeLessThan(dateAccuracy);
        expect(ws.getCell("C1").type).toEqual(Excel.ValueType.Date);
        
        if (checkFormulas) {
            expect(ws.getCell("D1").value).toEqual(utils.testValues.formulas[0]);
            expect(ws.getCell("D1").type).toEqual(Excel.ValueType.Formula);
            expect(ws.getCell("E1").value).toEqual(utils.testValues.formulas[1]);
            expect(ws.getCell("E1").type).toEqual(Excel.ValueType.Formula);
            expect(ws.getCell("F1").value).toEqual(utils.testValues.hyperlink);
            expect(ws.getCell("F1").type).toEqual(Excel.ValueType.Hyperlink);
            expect(ws.getCell("G1").value).toEqual(utils.testValues.str2);
        } else {
            expect(ws.getCell("D1").value).toEqual(utils.testValues.formulas[0].result);
            expect(ws.getCell("D1").type).toEqual(Excel.ValueType.Number);
            expect(ws.getCell("E1").value).toBeNull();
            expect(ws.getCell("E1").type).toEqual(Excel.ValueType.Null);
            expect(ws.getCell("F1").value).toEqual(utils.testValues.hyperlink.hyperlink);
            expect(ws.getCell("F1").type).toEqual(Excel.ValueType.String);
            expect(ws.getCell("G1").value).toEqual(utils.testValues.str2);
        }
        
        // A2:B3
        expect(ws.getCell("A2").value).toEqual(5);
        expect(ws.getCell("A2").type).toEqual(Excel.ValueType.Number);
        expect(ws.getCell("A2").master).toBe(ws.getCell("A2"));
        
        if (checkMerges) {
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
        }
        
        if (checkStyles) {
            expect(ws.getCell("A4").numFmt).toEqual(utils.testValues.numFmt1);
            expect(ws.getCell("A4").type).toEqual(Excel.ValueType.Number);
            expect(ws.getCell("A4").border).toEqual(utils.styles.borders.thin);
            expect(ws.getCell("C4").numFmt).toEqual(utils.testValues.numFmt2);
            expect(ws.getCell("C4").type).toEqual(Excel.ValueType.Number);
            expect(ws.getCell("C4").border).toEqual(utils.styles.borders.doubleRed);
            expect(ws.getCell("E4").border).toEqual(utils.styles.borders.thickRainbow);
            
            // test fonts and formats
            expect(ws.getCell("A5").value).toEqual(utils.testValues.str);
            expect(ws.getCell("A5").type).toEqual(Excel.ValueType.String);
            expect(ws.getCell("A5").font).toEqual(utils.styles.fonts.arialBlackUI14);
            expect(ws.getCell("B5").value).toEqual(utils.testValues.str);
            expect(ws.getCell("B5").type).toEqual(Excel.ValueType.String);
            expect(ws.getCell("B5").font).toEqual(utils.styles.fonts.broadwayRedOutline20);
            expect(ws.getCell("C5").value).toEqual(utils.testValues.str);
            expect(ws.getCell("C5").type).toEqual(Excel.ValueType.String);
            expect(ws.getCell("C5").font).toEqual(utils.styles.fonts.comicSansUdB16);
            
            expect(Math.abs(ws.getCell("D5").value - 1.6)).toBeLessThan(0.00000001);
            expect(ws.getCell("D5").type).toEqual(Excel.ValueType.Number);
            expect(ws.getCell("D5").numFmt).toEqual(utils.testValues.numFmt1);
            expect(ws.getCell("D5").font).toEqual(utils.styles.fonts.arialBlackUI14);
            
            expect(Math.abs(ws.getCell("E5").value - 1.6)).toBeLessThan(0.00000001);
            expect(ws.getCell("E5").type).toEqual(Excel.ValueType.Number);
            expect(ws.getCell("E5").numFmt).toEqual(utils.testValues.numFmt2);
            expect(ws.getCell("E5").font).toEqual(utils.styles.fonts.broadwayRedOutline20);
            
            expect(Math.abs(ws.getCell("F5").value.getTime() - utils.testValues.date.getTime())).toBeLessThan(dateAccuracy);
            expect(ws.getCell("F5").type).toEqual(Excel.ValueType.Date);
            expect(ws.getCell("F5").numFmt).toEqual(utils.testValues.numFmtDate);
            expect(ws.getCell("F5").font).toEqual(utils.styles.fonts.comicSansUdB16);
            
            expect(ws.getRow(5).height).not.toBeDefined();
            expect(ws.getRow(6).height).toEqual(42);
            _.each(utils.styles.alignments, function(alignment, index) {
                var rowNumber = 6
                var colNumber = index + 1;
                var cell = ws.getCell(rowNumber, colNumber);
                expect(cell.value).toEqual(alignment.text);
                expect(cell.alignment).toEqual(alignment.alignment);
            });
            
            if (checkBadAlignments) {
                _.each(utils.styles.badAlignments, function(alignment, index) {
                    var rowNumber = 7;
                    var colNumber = index + 1;
                    var cell = ws.getCell(rowNumber, colNumber);
                    expect(cell.value).toEqual(alignment.text);
                    expect(cell.alignment).not.toBeDefined();
                });
            }
            
            var row8 = ws.getRow(8);
            expect(row8.height).toEqual(40);
            expect(row8.getCell(1).fill).toEqual(utils.styles.fills.blueWhiteHGrad);
            expect(row8.getCell(2).fill).toEqual(utils.styles.fills.redDarkVertical);
            expect(row8.getCell(3).fill).toEqual(utils.styles.fills.redGreenDarkTrellis);
            expect(row8.getCell(4).fill).toEqual(utils.styles.fills.rgbPathGrad);
        }
    }
}