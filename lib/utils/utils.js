/**
 * Copyright (c) 2014 Guyon Roche
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 * 
 */
"use strict";

var fs = require('fs');
var Promise = require('bluebird');
var _ = require('underscore');

var moduleFile = {
    cache: {},
    
    read: function(path) {
        // if not already load-ed-ing...
        if (!this.cache[path]) {
            var deferred = Promise.defer();
            fs.readFile(path, 'utf8', function(error, data) {
                if (error) {
                    deferred.reject(err);
                } else {
                    deferred.resolve(data);
                }
            });
            this.cache[path] = deferred.promise;
        }
        return this.cache[path];
    }
};

var template = {
    cache: {},
    
    fetch: function(path) {
        // if not already load-ed-ing...
        if (!this.cache[path]) {
            this.cache[path] = moduleFile.read(path)
                .then(function(data) {
                    return _.template(data);
                });
        }
        return this.cache[path];
    }
};

// useful stuff
var inherits = function(cls, superCtor, statics, prototype) {
    cls.super_ = superCtor;
    
    if (!prototype) {
        prototype = statics;
        statics = null;
    }
    
    if (statics) {
        for (var i in statics) {
            Object.defineProperty(cls, i, Object.getOwnPropertyDescriptor(statics, i));
        }
    }
    
    var properties = {
        constructor: {
            value: cls,
            enumerable: false,
            writable: false,
            configurable: true
        }
    };
    if (prototype) {
        for (var i in prototype) {
            properties[i] = Object.getOwnPropertyDescriptor(prototype, i);
        }
    }
    
    cls.prototype = Object.create(superCtor.prototype, properties);
};


var utils = module.exports = {
    promiseImmediate: function(value) {
        var deferred = Promise.defer();
        if (global.setImmediate) {
            setImmediate(function() {
                deferred.resolve(value);
            });
        } else {
            // poorman's setImmediate - must wait at least 1ms
            setTimeout(function() {
                deferred.resolve(value);
            },1);
        }
        return deferred.promise;
    },
    readModuleFile: function(path) {
        return moduleFile.read(path);
    },
    fetchTemplate: function(path) {
        return template.fetch(path);
    },
    inherits: inherits,
    dateToExcel: function(d) {
        return 25569 + d.getTime() / (24 * 3600 * 1000);
    },
    excelToDate: function(v) {
        return new Date((v - 25569) * 24 * 3600 * 1000);
    },
    parsePath: function(filepath) {
        var last = filepath.lastIndexOf('/');
        return {
            path: filepath.substring(0, last),
            name: filepath.substring(last + 1)
        };
    },
    getRelsPath: function(filepath) {
        var path = utils.parsePath(filepath);
        return path.path + '/_rels/' + path.name + '.rels'; 
    },
    xmlEncode: function(text) {
        return text.replace(/[<>&'"]/g, function (c) {
            switch (c) {
                case '<': return '&lt;';
                case '>': return '&gt;';
                case '&': return '&amp;';
                case '\'': return '&apos;';
                case '"': return '&quot;';
                default: return '';
            }
        });
    },
    xmlDecode: function(text) {
        return text.replace(/&([a-z]*);/, function(c) {
            switch (c) {
                case '&lt;': return '<';
                case '&gt;': return '>';
                case '&amp;': return '&';
                case '&apos;': return '\'';
                case '&quot;': return '"';
                default: return c;
            }
        });
    },
    validInt: function(value) {
        var i = parseInt(value);
        return i !== NaN ? i : 0;
    },
    
    fs: {
        exists: function(path) {
            var deferred = Promise.defer();
            fs.exists(path, function(exists) {
                deferred.resolve(exists);
            });
            return deferred.promise;
        }
    }

};
