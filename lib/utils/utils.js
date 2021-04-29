const fs = require('fs');

// useful stuff
const inherits = function(cls, superCtor, statics, prototype) {
  // eslint-disable-next-line no-underscore-dangle
  cls.super_ = superCtor;

  if (!prototype) {
    prototype = statics;
    statics = null;
  }

  if (statics) {
    Object.keys(statics).forEach(i => {
      Object.defineProperty(cls, i, Object.getOwnPropertyDescriptor(statics, i));
    });
  }

  const properties = {
    constructor: {
      value: cls,
      enumerable: false,
      writable: false,
      configurable: true,
    },
  };
  if (prototype) {
    Object.keys(prototype).forEach(i => {
      properties[i] = Object.getOwnPropertyDescriptor(prototype, i);
    });
  }

  cls.prototype = Object.create(superCtor.prototype, properties);
};

// eslint-disable-next-line no-control-regex
const xmlDecodeRegex = /[<>&'"\x7F\x00-\x08\x0B-\x0C\x0E-\x1F]/;
const utils = {
  nop() {},
  promiseImmediate(value) {
    return new Promise(resolve => {
      if (global.setImmediate) {
        setImmediate(() => {
          resolve(value);
        });
      } else {
        // poorman's setImmediate - must wait at least 1ms
        setTimeout(() => {
          resolve(value);
        }, 1);
      }
    });
  },
  inherits,
  dateToExcel(d, date1904) {
    return 25569 + ( d.getTime() / (24 * 3600 * 1000) ) - (date1904 ? 1462 : 0);
  },
  excelToDate(v, date1904) {
    const millisecondSinceEpoch = Math.round((v - 25569 + (date1904 ? 1462 : 0)) * 24 * 3600 * 1000);
    return new Date(millisecondSinceEpoch);
  },
  parsePath(filepath) {
    const last = filepath.lastIndexOf('/');
    return {
      path: filepath.substring(0, last),
      name: filepath.substring(last + 1),
    };
  },
  getRelsPath(filepath) {
    const path = utils.parsePath(filepath);
    return `${path.path}/_rels/${path.name}.rels`;
  },
  xmlEncode(text) {
    const regexResult = xmlDecodeRegex.exec(text);
    if (!regexResult) return text;

    let result = '';
    let escape = '';
    let lastIndex = 0;
    let i = regexResult.index;
    for (; i < text.length; i++) {
      const charCode = text.charCodeAt(i);
      switch (charCode) {
        case 34: // "
          escape = '&quot;';
          break;
        case 38: // &
          escape = '&amp;';
          break;
        case 39: // '
          escape = '&apos;';
          break;
        case 60: // <
          escape = '&lt;';
          break;
        case 62: // >
          escape = '&gt;';
          break;
        case 127:
          escape = '';
          break;
        default: {
          if (charCode <= 31 && (charCode <= 8 || (charCode >= 11 && charCode !== 13))) {
            escape = '';
            break;
          }
          continue;
        }
      }
      if (lastIndex !== i) result += text.substring(lastIndex, i);
      lastIndex = i + 1;
      if (escape) result += escape;
    }
    if (lastIndex !== i) return result + text.substring(lastIndex, i);
    return result;
  },
  xmlDecode(text) {
    return text.replace(/&([a-z]*);/g, c => {
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
  validInt(value) {
    const i = parseInt(value, 10);
    return !Number.isNaN(i) ? i : 0;
  },

  isDateFmt(fmt) {
    if (!fmt) {
      return false;
    }

    // must remove all chars inside quotes and []
    fmt = fmt.replace(/\[[^\]]*]/g, '');
    fmt = fmt.replace(/"[^"]*"/g, '');
    // then check for date formatting chars
    const result = fmt.match(/[ymdhMsb]+/) !== null;
    return result;
  },

  fs: {
    exists(path) {
      return new Promise(resolve => {
        fs.access(path, fs.constants.F_OK, err => {
          resolve(!err);
        });
      });
    },
  },

  toIsoDateString(dt) {
    return dt.toIsoString().subsstr(0, 10);
  },
};

module.exports = utils;
