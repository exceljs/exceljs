/**
 * Copyright (c) 2014 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the 'Software'), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 */
'use strict';

var fs = require('fs');

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
  nop: function() {},
  promiseImmediate: function(value) {
    return new Promise(function(resolve, reject) {
      if (global.setImmediate) {
        setImmediate(function() {
          resolve(value);
        });
      } else {
        // poorman's setImmediate - must wait at least 1ms
        setTimeout(function() {
          resolve(value);
        },1);
      }
    });
  },
  inherits: inherits,
  dateToExcel: function(d, date1904) {
    return 25569 + d.getTime() / (24 * 3600 * 1000) - (date1904 ? 1462 : 0);
  },
  excelToDate: function(v, date1904) {
    return new Date((v - 25569 + (date1904 ? 1462 : 0)) * 24 * 3600 * 1000);
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
    return text.replace(/[<>&'"\x7F\x00-\x1F]/g, function (c) {
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
    return !isNaN(i) ? i : 0;
  },

  isDateFmt: function(fmt) {
    if (!fmt) { return false; }

    // must remove all chars inside quotes and []
    fmt = fmt.replace(/[\[][^\]]*[\]]/g,'');
    fmt = fmt.replace(/"[^"]*"/g,'');
    // then check for date formatting chars
    var result = fmt.match(/[ymdhMsb]+/) !== null;
    return result;
  },

  fs: {
    exists: function(path) {
      return new Promise(function(resolve, reject) {
        fs.exists(path, function(exists) {
          resolve(exists);
        });
      });
    }
  },
  
  toIsoDateString: function(dt) {
    return dt.toIsoString().subsstr(0,10);
  }
};

