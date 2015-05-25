var _ = require("underscore");

var StreamBuf = require("../lib/utils/stream-buf");
var StringBuf = require("../lib/utils/string-buf");

describe("StreamBuf", function() {
    // StreamBuf is designed as a general-purpose writable-readable stream
    // However its use in ExcelJS is primarily as a memory buffer between
    // the streaming writers and the archive, hence the tests here will
    // focus only on that.
    it("writes strings as UTF8", function() {
        var stream = new StreamBuf();
        stream.write("Hello, World!");
        var chunk = stream.read();
        expect(chunk instanceof Buffer).toBeTruthy();
        expect(chunk.toString("UTF8")).toEqual("Hello, World!");
    });
    
    it("writes StringBuf chunks", function() {
        var stream = new StreamBuf();
        var strBuf = new StringBuf({size: 64});
        strBuf.addText("Hello, World!");
        stream.write(strBuf);
        var chunk = stream.read();
        expect(chunk instanceof Buffer).toBeTruthy();
        expect(chunk.toString("UTF8")).toEqual("Hello, World!");
    });
    
    it("signals end", function(done) {
        var stream = new StreamBuf();
        stream.on("end", function() {
            done();
        });
        stream.write("Hello, World!");
        stream.end();
    });
    
});
