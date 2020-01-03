var fs = require('fs');
var _ = require('underscore');
var Promise = require('bluebird');

var main = module.exports = {
    cleanDir: function(path) {
        var deferred = Promise.defer();
        
        var remove = function(file) {
            var myDeferred = Promise.defer();
            var myHandler = function(err) { if (err) { myDeferred.reject(err)} else { myDeferred.resolve(); }}
            fs.stat(file, function(err, stat) {
                if (err) {
                    myDeferred.reject(err);
                } else {
                    if (stat.isFile()) {
                        console.log("unlink " + file);
                        fs.unlink(file, myHandler);
                    } else if (stat.isDirectory()) {
                        main.cleanDir(file)
                            .then(function() {
                                console.log("rmdir " + file);
                                fs.rmdir(file, myHandler);
                            })
                            .catch(myHandler);
                    }
                }
            });
            return myDeferred.promise;
        }
        
        fs.readdir(path, function(err, files) {
            if(err) {
                deferred.reject(err);
            } else {
                var promises = [];
                _.each(files, function(file) {
                    promises.push(remove(path + "/" + file));
                });
                
                Promise.all(promises)
                    .then(function() {
                        deferred.resolve();
                    })
                    .catch(function(err) {
                        deferred.reject(err);
                    });
            }
        });
        
        return deferred.promise;
    },
    
    randomName: function(length) {
        length = length || 5;
        var text = [];
        var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    
        for( var i=0; i < length; i++ )
            text.push(possible.charAt(Math.floor(Math.random() * possible.length)));
    
        return text.join('');
    },
    randomNum: function(d) {
        return Math.round(Math.random()*d);
    },
    
    fmt: {
        number: function(n) {
            // output large numbers with thousands separator
            var s = n.toString();
            var l = s.length;
            var a = [];
            var r = (l % 3) || 3;
            var i = 0;
            while (i < l) {
                a.push(s.substr(i,r));
                i += r;
                r = 3;
            }
            return a.join(',');
        } 
    }

};