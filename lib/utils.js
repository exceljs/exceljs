/**
 * Copyright (c) 2014 Guyon Roche
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:</p>
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
    errors: {},
    promises: {},
    
    read: function(path) {
        var self = this;
        var filename = require.resolve(path);
        
        // if already loaded...
        if (this.cache[filename]) {
            return Promise.resolve(this.cache[filename]);
        }
        
        // if already failed...
        if (this.errors[filename]) {
            return Promise.reject(this.errors[filename]);
        }
        
        // if currently being loaded, tag along...
        if (this.promises[filename]) {
            return this.promises[filename];
        }
        
        var deferred = this.promises[filename] = Promise.defer();
        
        fs.readFile(filename, 'utf8', function(err, data) {
            delete self.promises[filename];
            if (err) {
                self.errors[filename] = err;
                deferred.reject(err);
            } else {
                self.cache[filename] = data;
                deferred.resolve(data);
            }
        });
        
        return deferred.promise;
    }
};

var template = {
    cache: {},
    errors: {},
    promises: {},
    
    _template: null,
    _templatePromise: null,
    fetch: function(path) {
        var self = this;

        // if already fetched, job done...
        if (this.cache[path]) {
            return Promise.resolve(this.cache[path]);
        }
        
        // if currently fetching, tag along...
        if (this.promises[path]) {
            return this.promises[path];
        }
        
        // nope, it's up to me...
        return this.promises[path] = moduleFile.read(path)
            .then(function(data) {
                return self.cache[path] = _.template(data);
            });
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
    }
};
