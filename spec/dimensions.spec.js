var _ = require("underscore");
var Dimensions = require("../lib/dimensions");

describe("Dimensions", function() {

    it("has a valid default value", function() {
        var d = new Dimensions();
        
        expect(d.tl).toEqual("A1");
        expect(d.br).toEqual("A1");
        expect(d.toString()).toEqual("A1:A1");
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