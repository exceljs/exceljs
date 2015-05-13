var _ = require("underscore");
var colCache = require("../lib/utils/col-cache");

describe("colCache", function() {

    it("caches values", function() {
        expect(colCache.l2n("A")).toEqual(1);
        expect(colCache._l2n.A).toEqual(1);
        expect(colCache._n2l[1]).toEqual("A");
        
        // also, because of the fill heuristic A-Z will be there too
        var dic = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];
        _.each(dic, function(letter, index) {
            expect(colCache._l2n[letter]).toEqual(index+1);
            expect(colCache._n2l[index+1]).toEqual(letter);
        });
        
        // next level
        expect(colCache.n2l(27)).toEqual("AA");
        expect(colCache._l2n.AB).toEqual(28);
        expect(colCache._n2l[28]).toEqual("AB");
    });
    
    it("converts numbers to letters", function() {
        expect(colCache.n2l(1)).toEqual("A");
        expect(colCache.n2l(26)).toEqual("Z");
        expect(colCache.n2l(27)).toEqual("AA");
        expect(colCache.n2l(702)).toEqual("ZZ");
        expect(colCache.n2l(703)).toEqual("AAA");
    });
    it("converts letters to numbers", function() {
        expect(colCache.l2n("A")).toEqual(1);
        expect(colCache.l2n("Z")).toEqual(26);
        expect(colCache.l2n("AA")).toEqual(27);
        expect(colCache.l2n("ZZ")).toEqual(702);
        expect(colCache.l2n("AAA")).toEqual(703);
    });

    it("throws when out of bounds", function() {
        expect(function() { colCache.n2l(0); }).toThrow();
        expect(function() { colCache.n2l(-1); }).toThrow();
        expect(function() { colCache.n2l(16385); }).toThrow();

        expect(function() { colCache.l2n(""); }).toThrow();
        expect(function() { colCache.l2n("AAAA"); }).toThrow();
        expect(function() { colCache.l2n(16385); }).toThrow();
    });
    
    it("validates addresses properly", function() {
        expect(colCache.validateAddress("A1")).toBeTruthy();
        expect(colCache.validateAddress("AA10")).toBeTruthy();
        expect(colCache.validateAddress("ABC100000")).toBeTruthy();
        
        expect(function() { colCache.validateAddress("A"); }).toThrow();
        expect(function() { colCache.validateAddress("1"); }).toThrow();
        expect(function() { colCache.validateAddress("1A"); }).toThrow();
        expect(function() { colCache.validateAddress("A 1"); }).toThrow();
        expect(function() { colCache.validateAddress("A1A"); }).toThrow();
        expect(function() { colCache.validateAddress("1A1"); }).toThrow();
        expect(function() { colCache.validateAddress("a1"); }).toThrow();
        expect(function() { colCache.validateAddress("a"); }).toThrow();
    });
    
    it("decodes addresses", function() {
        expect(colCache.decodeAddress("A1")).toEqual({address:"A1",col:1, row: 1});
        expect(colCache.decodeAddress("AA11")).toEqual({address:"AA11",col:27, row: 11});
    });
    
    it("gets address structures (and caches them)", function() {
        var addr = colCache.getAddress("D5");
        expect(addr.address).toEqual("D5");
        expect(addr.row).toEqual(5);
        expect(addr.col).toEqual(4);
        expect(colCache.getAddress("D5")).toBe(addr);
        expect(colCache.getAddress(5,4)).toBe(addr);
        
        addr = colCache.getAddress("E4");
        expect(addr.address).toEqual("E4");
        expect(addr.row).toEqual(4);
        expect(addr.col).toEqual(5);
        expect(colCache.getAddress("E4")).toBe(addr);
        expect(colCache.getAddress(4,5)).toBe(addr);
    });
    
    it("decodes addresses and ranges", function() {
        // address
        expect(colCache.decode("A1")).toEqual({address:"A1",col:1, row: 1});
        expect(colCache.decode("AA11")).toEqual({address:"AA11",col:27, row: 11});
        
        // range
        expect(colCache.decode("A1:B2")).toEqual({
            dimensions: "A1:B2",
            tl: "A1",
            br: "B2",
            top: 1,
            left: 1,
            bottom: 2,
            right: 2
        });
        // wonky ranges
        expect(colCache.decode("A2:B1")).toEqual({
            dimensions: "A1:B2",
            tl: "A1",
            br: "B2",
            top: 1,
            left: 1,
            bottom: 2,
            right: 2
        });
        expect(colCache.decode("B2:A1")).toEqual({
            dimensions: "A1:B2",
            tl: "A1",
            br: "B2",
            top: 1,
            left: 1,
            bottom: 2,
            right: 2
        });
    });
  });