var Promise = require("bluebird");
var MemoryStream = require("memorystream");

var Excel = require("../excel");

var main = module.exports = {
    cloneByModel: function(thing1, type) {
        var thing2 = new type();
        thing2.model = thing1.model;
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
    }
}