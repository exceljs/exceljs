var _ = require("underscore");
var Dimensions = require("../lib/dimensions");

describe("Dimensions", function() {
    
    function check(d, range, tl, br, top, left, bottom, right) {
        expect(d.range).toEqual(range);
        expect(d.tl).toEqual(tl);
        expect(d.br).toEqual(br);
        expect(d.top).toEqual(top);
        expect(d.left).toEqual(left);
        expect(d.bottom).toEqual(bottom);
        expect(d.right).toEqual(right);
        expect(d.toString()).toEqual(range);
    }
    
    it("has a valid default value", function() {
        var d = new Dimensions();
        check(d, "A1:A1", "A1", "A1", 1, 1, 1, 1);
    });
    
    it("constructs as expected", function() {
        // check range + rotations
        check(new Dimensions("B5:D10"), "B5:D10", "B5", "D10", 5, 2, 10, 4);
        check(new Dimensions("B10:D5"), "B5:D10", "B5", "D10", 5, 2, 10, 4);
        check(new Dimensions("D5:B10"), "B5:D10", "B5", "D10", 5, 2, 10, 4);
        check(new Dimensions("D10:B5"), "B5:D10", "B5", "D10", 5, 2, 10, 4);
        
        check(new Dimensions("C7","G16"), "C7:G16", "C7", "G16", 7, 3, 16, 7);
        check(new Dimensions("C16","G7"), "C7:G16", "C7", "G16", 7, 3, 16, 7);
        check(new Dimensions("G7","C16"), "C7:G16", "C7", "G16", 7, 3, 16, 7);
        check(new Dimensions("G16","C7"), "C7:G16", "C7", "G16", 7, 3, 16, 7);
        
        check(new Dimensions(7, 3, 16, 7), "C7:G16", "C7", "G16", 7, 3, 16, 7);
        check(new Dimensions(16, 3, 7, 7), "C7:G16", "C7", "G16", 7, 3, 16, 7);
        check(new Dimensions(7, 7, 16, 3), "C7:G16", "C7", "G16", 7, 3, 16, 7);
        check(new Dimensions(16, 7, 7, 3), "C7:G16", "C7", "G16", 7, 3, 16, 7);

        check(new Dimensions([7, 3, 16, 7]), "C7:G16", "C7", "G16", 7, 3, 16, 7);
        check(new Dimensions([16, 3, 7, 7]), "C7:G16", "C7", "G16", 7, 3, 16, 7);
        check(new Dimensions([7, 7, 16, 3]), "C7:G16", "C7", "G16", 7, 3, 16, 7);
        check(new Dimensions([16, 7, 7, 3]), "C7:G16", "C7", "G16", 7, 3, 16, 7);

        check(new Dimensions("B5"), "B5:B5", "B5", "B5", 5, 2, 5, 2);
    });
    
    
    it("expands properly", function() {
        var d = new Dimensions();
        
        d.expand(1,1,1,3);
        expect(d.tl).toEqual("A1");
        expect(d.br).toEqual("C1");
        expect(d.toString()).toEqual("A1:C1");
        
        d.expand(1,3,3,3);
        expect(d.tl).toEqual("A1");
        expect(d.br).toEqual("C3");
        expect(d.toString()).toEqual("A1:C3");
    });
    
    it("doesn't always include the default row/col", function() {
        var d = new Dimensions();
        
        d.expand(2,2,4,4);
        expect(d.tl).toEqual("B2");
        expect(d.br).toEqual("D4");
        expect(d.toString()).toEqual("B2:D4");
    });
});