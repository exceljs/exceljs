var fs = require("fs");
var _ = require("underscore");
var Excel = require("../excel");
var utils = require("./testutils");

describe("WorkbookWriter", function() {
       
    xit("creates sheets with correct names", function() {
        var wb = new Excel.stream.xlsx.WorkbookWriter();
        var ws1 = wb.addWorksheet("Hello, World!");
        expect(ws1.name).toEqual("Hello, World!");
        
        var ws2 = wb.addWorksheet();
        expect(ws2.name).toMatch(/sheet\d+/);
    });
    
    xit("serializes to xlsx file properly", function(done) {
        
        var wb = utils.createTestBook(true, Excel.stream.xlsx.WorkbookWriter, {filename: "./wbw.test.xlsx"});
        //fs.writeFileSync("./testmodel.json", JSON.stringify(wb.model, null, "    "));
        
        wb.commit()
            .then(function() {
                var wb2 = new Excel.Workbook();
                return wb2.xlsx.readFile("./wb.test.xlsx");
            })
            .then(function(wb2) {
                checkTestBook(wb2, "xlsx");
            })
            .finally(function() {
                fs.unlink("./wb.test.xlsx", function(error) {
                    expect(error && error.message).toBeFalsy();
                    done();
                });
            });
    });
    
    it("serializes row styles and columns properly", function(done) {
        var options = {
            filename: "./wbw.test.xlsx",
            useStyles: true
        };
        var wb = new Excel.stream.xlsx.WorkbookWriter(options);
        var ws = wb.addWorksheet("blort");
        
        var colStyle = {
            font: utils.styles.fonts.comicSansUdB16,
            alignment: utils.styles.namedAlignments.middleCentre
        };
        ws.columns = [
            { header: "A1", width: 10 },
            { header: "B1", width: 20, style: colStyle },
            { header: "C1", width: 30 },
        ];
        
        ws.getRow(2).font = utils.styles.fonts.broadwayRedOutline20;
        
        ws.getCell("A2").value = "A2";
        ws.getCell("B2").value = "B2";
        ws.getCell("C2").value = "C2";
        ws.getCell("A3").value = "A3";
        ws.getCell("B3").value = "B3";
        ws.getCell("C3").value = "C3";
        
        wb.commit().delay(100)
            .then(function() {
                var wb2 = new Excel.Workbook();
                return wb2.xlsx.readFile("./wbw.test.xlsx");
            })
            .then(function(wb2) {
                var ws2 = wb2.getWorksheet("blort");
                _.each(["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"], function(address) {
                    expect(ws2.getCell(address).value).toEqual(address);
                });
                expect(ws2.getCell("B1").font).toEqual(utils.styles.fonts.comicSansUdB16);
                expect(ws2.getCell("B1").alignment).toEqual(utils.styles.namedAlignments.middleCentre);
                expect(ws2.getCell("A2").font).toEqual(utils.styles.fonts.broadwayRedOutline20);
                expect(ws2.getCell("B2").font).toEqual(utils.styles.fonts.broadwayRedOutline20);
                expect(ws2.getCell("C2").font).toEqual(utils.styles.fonts.broadwayRedOutline20);                
                expect(ws2.getCell("B3").font).toEqual(utils.styles.fonts.comicSansUdB16);
                expect(ws2.getCell("B3").alignment).toEqual(utils.styles.namedAlignments.middleCentre);
                
                expect(ws2.getColumn(2).font).toEqual(utils.styles.fonts.comicSansUdB16);
                expect(ws2.getColumn(2).alignment).toEqual(utils.styles.namedAlignments.middleCentre);
                
                expect(ws2.getRow(2).font).toEqual(utils.styles.fonts.broadwayRedOutline20);
            })
            .finally(function() {
                //fs.unlink("./wbw.test.xlsx", function(error) {
                //    expect(error && error.message).toBeFalsy();
                //    done();
                //});
                done();
            });
    });
    
    xit("serializes and deserializes a lot of sheets to xlsx file properly", function(done) {
        var wb = new Excel.stream.xlsx.WorkbookWriter({filename: "./wbw.test.xlsx"});
        var numSheets = 90;
        // add numSheets sheets
        for (i = 1; i <= numSheets; i++) {
            var ws = wb.addWorksheet("sheet" + i);
            ws.getCell("A1").value = i;
        }
        wb.xlsx.writeFile("./wbw.test.xlsx")
            .then(function() {
                var wb2 = new Excel.Workbook();
                return wb2.xlsx.readFile("./wbw.test.xlsx");
            })
            .then(function(wb2) {
                for (i = 1; i <= numSheets; i++) {
                    var ws = wb.getWorksheet("sheet" + i);
                    expect(ws).toBeDefined();
                    expect(ws.getCell("A1").value).toEqual(i);
                }
            })
            .finally(function() {
                fs.unlink("./wbw.test.xlsx", function(error) {
                    expect(error && error.message).toBeFalsy();
                    done();
                });
            });
    });
});
