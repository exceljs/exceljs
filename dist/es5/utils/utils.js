/**
 * Copyright (c) 2014-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var fs = require('fs');
var PromishLib = require('./promish');

// useful stuff
var inherits = function inherits(cls, superCtor, statics, prototype) {
  // eslint-disable-next-line no-underscore-dangle
  cls.super_ = superCtor;

  if (!prototype) {
    prototype = statics;
    statics = null;
  }

  if (statics) {
    Object.keys(statics).forEach(function (i) {
      Object.defineProperty(cls, i, Object.getOwnPropertyDescriptor(statics, i));
    });
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
    Object.keys(prototype).forEach(function (i) {
      properties[i] = Object.getOwnPropertyDescriptor(prototype, i);
    });
  }

  cls.prototype = Object.create(superCtor.prototype, properties);
};

var utils = module.exports = {
  nop: function nop() {},
  promiseImmediate: function promiseImmediate(value) {
    return new PromishLib.Promish(function (resolve) {
      if (global.setImmediate) {
        setImmediate(function () {
          resolve(value);
        });
      } else {
        // poorman's setImmediate - must wait at least 1ms
        setTimeout(function () {
          resolve(value);
        }, 1);
      }
    });
  },
  inherits: inherits,
  dateToExcel: function dateToExcel(d, date1904) {
    return 25569 + d.getTime() / (24 * 3600 * 1000) - (date1904 ? 1462 : 0);
  },
  excelToDate: function excelToDate(v, date1904) {
    var millisecondSinceEpoch = Math.round((v - 25569 + (date1904 ? 1462 : 0)) * 24 * 3600 * 1000);
    return new Date(millisecondSinceEpoch);
  },
  parsePath: function parsePath(filepath) {
    var last = filepath.lastIndexOf('/');
    return {
      path: filepath.substring(0, last),
      name: filepath.substring(last + 1)
    };
  },
  getRelsPath: function getRelsPath(filepath) {
    var path = utils.parsePath(filepath);
    return path.path + '/_rels/' + path.name + '.rels';
  },
  xmlEncode: function xmlEncode(text) {
    // eslint-disable-next-line no-control-regex
    return text.replace(/[<>&'"\x7F\x00-\x08\x0B-\x0C\x0E-\x1F]/g, function (c) {
      switch (c) {
        case '<':
          return '&lt;';
        case '>':
          return '&gt;';
        case '&':
          return '&amp;';
        case '\'':
          return '&apos;';
        case '"':
          return '&quot;';
        default:
          return '';
      }
    });
  },
  xmlDecode: function xmlDecode(text) {
    return text.replace(/&([a-z]*);/, function (c) {
      switch (c) {
        case '&lt;':
          return '<';
        case '&gt;':
          return '>';
        case '&amp;':
          return '&';
        case '&apos;':
          return '\'';
        case '&quot;':
          return '"';
        default:
          return c;
      }
    });
  },
  validInt: function validInt(value) {
    var i = parseInt(value, 10);
    return !isNaN(i) ? i : 0;
  },

  isDateFmt: function isDateFmt(fmt) {
    if (!fmt) {
      return false;
    }

    // must remove all chars inside quotes and []
    fmt = fmt.replace(/\[[^\]]*]/g, '');
    fmt = fmt.replace(/"[^"]*"/g, '');
    // then check for date formatting chars
    var result = fmt.match(/[ymdhMsb]+/) !== null;
    return result;
  },

  fs: {
    exists: function exists(path) {
      return new PromishLib.Promish(function (resolve) {
        fs.exists(path, function (exists) {
          resolve(exists);
        });
      });
    }
  },

  toIsoDateString: function toIsoDateString(dt) {
    return dt.toIsoString().subsstr(0, 10);
  }
};
//# sourceMappingURL=utils.js.map
