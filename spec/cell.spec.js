var _ = require("underscore");
var Cell = require("../lib/cell");
var Enums = require("../lib/enums");

function createSheetMock() {
    return {
        cells: {},
        getCell: function(address) {
            return this.cells[address];
        }
    };
}

describe("Cell", function() {
    it("stores values", function() {
        var sheet = createSheetMock();        
        
        sheet.cells.A1 = new Cell({address:"A1",worksheet:sheet});
        expect(sheet.cells.A1.type).toEqual(Enums.ValueType.Null);
        
        expect(sheet.cells.A1.value = 5).toEqual(5);
        expect(sheet.cells.A1.value).toEqual(5);
        expect(sheet.cells.A1.type).toEqual(Enums.ValueType.Number);
        
        var strValue = "Hello, World!";
        expect(sheet.cells.A1.value = strValue).toEqual(strValue);
        expect(sheet.cells.A1.value).toEqual(strValue);
        expect(sheet.cells.A1.type).toEqual(Enums.ValueType.String);
        
        var dateValue = new Date();
        expect(sheet.cells.A1.value = dateValue).toEqual(dateValue);
        expect(sheet.cells.A1.value).toEqual(dateValue);
        expect(sheet.cells.A1.type).toEqual(Enums.ValueType.Date);
        
        var formulaValue = {formula: "A2", result: 5};
        expect(sheet.cells.A1.value = formulaValue).toEqual(formulaValue);
        expect(sheet.cells.A1.value).toEqual(formulaValue);
        expect(sheet.cells.A1.type).toEqual(Enums.ValueType.Formula);
        
        // no result
        formulaValue = {formula: "A3"};
        expect(sheet.cells.A1.value = formulaValue).toEqual(formulaValue);
        expect(sheet.cells.A1.value).toEqual(formulaValue);
        expect(sheet.cells.A1.type).toEqual(Enums.ValueType.Formula);
        
        var hyperlinkValue = {hyperlink: "http://www.link.com", text: "www.link.com"};
        expect(sheet.cells.A1.value = hyperlinkValue).toEqual(hyperlinkValue);
        expect(sheet.cells.A1.value).toEqual(hyperlinkValue);
        expect(sheet.cells.A1.type).toEqual(Enums.ValueType.Hyperlink);
        
        expect(sheet.cells.A1.value = null).toEqual(null);
        expect(sheet.cells.A1.type).toEqual(Enums.ValueType.Null);
    });
    it("validates options on construction", function() {
        var sheet = createSheetMock();
        expect(function() { new Cell1(); }).toThrow();
        expect(function() { new Cell1({worksheet: sheet}); }).toThrow();
        expect(function() { new Cell1({address:"A", worksheet: sheet}); }).toThrow();
        expect(function() { new Cell1({address:"Hello, World!", worksheet: sheet}); }).toThrow();
        expect(function() { new Cell1({address:"A1"}); }).toThrow();
    });
    it("merges", function() {
        var sheet = createSheetMock();        
        
        sheet.cells.A1 = new Cell({address:"A1",worksheet:sheet});
        sheet.cells.A2 = new Cell({address:"A2",worksheet:sheet});
        
        sheet.cells.A1.value = 5;
        sheet.cells.A2.value = "Hello, World!";
        
        sheet.cells.A2.merge(sheet.cells.A1);
        
        expect(sheet.cells.A2.value).toEqual(5);
        expect(sheet.cells.A2.type).toEqual(Enums.ValueType.Merge);
        expect(sheet.cells.A1._mergeCount).toEqual(1);
        expect(sheet.cells.A1.isMerged).toBeTruthy();
        expect(sheet.cells.A2.isMerged).toBeTruthy();
        expect(sheet.cells.A2.isMergedTo(sheet.cells.A1)).toBeTruthy();
        expect(sheet.cells.A2.master).toBe(sheet.cells.A1);
        expect(sheet.cells.A1.master).toBe(sheet.cells.A1);
        
        // assignment of slaves write to the master
        sheet.cells.A2.value = 7;
        expect(sheet.cells.A1.value).toEqual(7);
        
        // assignment of strings should add 1 ref
        var strValue = "Boo!";
        sheet.cells.A2.value = strValue;
        expect(sheet.cells.A1.value).toEqual(strValue);
        
        // unmerge should work also
        sheet.cells.A2.unmerge();
        expect(sheet.cells.A2.type).toEqual(Enums.ValueType.Null);
        expect(sheet.cells.A1._mergeCount).toEqual(0);
        expect(sheet.cells.A1.isMerged).toBeFalsy();
        expect(sheet.cells.A2.isMerged).toBeFalsy();
        expect(sheet.cells.A2.isMergedTo(sheet.cells.A1)).toBeFalsy();
        expect(sheet.cells.A2.master).toBe(sheet.cells.A2);
        expect(sheet.cells.A1.master).toBe(sheet.cells.A1);
    });
    
    it("upgrades from string to hyperlink", function() {
        var sheet = createSheetMock();
        
        var strValue = "www.link.com";
        var linkValue = "http://www.link.com";
        
        sheet.cells.A1 = new Cell({address:"A1",worksheet:sheet});
        sheet.cells.A1.value = strValue;
        
        sheet.cells.A1._upgradeToHyperlink(linkValue);
        
        expect(sheet.cells.A1.type).toEqual(Enums.ValueType.Hyperlink);
    });
    
    it("doesn't upgrade from non-string to hyperlink", function() {
        var sheet = createSheetMock();        
        
        var linkValue = "http://www.link.com";

        sheet.cells.A1 = new Cell({address:"A1",worksheet:sheet});
        
        // null
        sheet.cells.A1._upgradeToHyperlink(linkValue);
        expect(sheet.cells.A1.type).toEqual(Enums.ValueType.Null);
        
        // number
        sheet.cells.A1.value = 5;
        sheet.cells.A1._upgradeToHyperlink(linkValue);
        expect(sheet.cells.A1.type).toEqual(Enums.ValueType.Number);
        
        // date
        sheet.cells.A1.value = new Date();
        sheet.cells.A1._upgradeToHyperlink(linkValue);
        expect(sheet.cells.A1.type).toEqual(Enums.ValueType.Date);
        
        // formula
        sheet.cells.A1.value = {formula: "A2"};
        sheet.cells.A1._upgradeToHyperlink(linkValue);
        expect(sheet.cells.A1.type).toEqual(Enums.ValueType.Formula);
        
        // hyperlink
        sheet.cells.A1.value = {hyperlink: "http://www.link2.com", text: "www.link2.com"};
        sheet.cells.A1._upgradeToHyperlink(linkValue);
        expect(sheet.cells.A1.type).toEqual(Enums.ValueType.Hyperlink);
        
        // cleanup
        sheet.cells.A1.value = null;
    });
});