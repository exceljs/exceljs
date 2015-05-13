var _ = require("underscore");

var SharedStrings = require("../lib/utils/shared-strings");

describe("SharedStrings", function() {
    
    it("Stores and shares string values", function() {
        var ss = new SharedStrings();
        
        var sHello = ss.add("Hello");
        var sHello_v2 = ss.add("Hello");
        var sGoodbye = ss.add("Goodbye");
        
        expect(sHello).toEqual(sHello_v2);
        expect(sGoodbye).not.toEqual(sHello_v2);
        
        expect(ss.count).toEqual(2);
        expect(ss.totalRefs).toEqual(3);
    });    
});
