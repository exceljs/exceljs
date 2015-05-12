var _ = require("underscore");
var utils = require("../lib/utils/utils");

describe("utils", function() {
    it("encodes xml text", function() {
        
        expect(utils.xmlEncode('<')).toEqual('&lt;');
        expect(utils.xmlEncode('>')).toEqual('&gt;');
        expect(utils.xmlEncode('&')).toEqual('&amp;');
        expect(utils.xmlEncode('"')).toEqual('&quot;');
        expect(utils.xmlEncode("'")).toEqual('&apos;');
        
        expect(utils.xmlEncode('<a href="www.whatever.com">Talk to the H&</a>')).toEqual('&lt;a href=&quot;www.whatever.com&quot;&gt;Talk to the H&amp;&lt;/a&gt;');
    });
});