var Promise = require("bluebird");

var promise = Promise.resolve(1)
    .then(function(value) {
        return value + 1;
    });
    
setTimeout(function() {
    promise.then(function(value) {
        console.log("Got value: " + value);
    });
}, 2000);