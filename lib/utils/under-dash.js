const {toString} = Object.prototype;
const escapeHtmlRegex = /["&<>]/;

// [exceljs-hardened] CVE fix — prototype pollution (GHSA pending).
// deepMerge() used to walk every own key of the source object, including
// "__proto__", "constructor" and "prototype", and assign straight into
// whatever it resolved to. A JSON-parsed object can carry "__proto__" as
// an ordinary own enumerable key (unlike an object literal), so an
// attacker-controlled value merged this way could reach and mutate the
// real global Object.prototype. These keys are never legitimate merge
// targets, so they are skipped unconditionally.
const UNSAFE_MERGE_KEYS = new Set(['__proto__', 'constructor', 'prototype']);

const _ = {
  each: function each(obj, cb) {
    if (obj) {
      if (Array.isArray(obj)) {
        obj.forEach(cb);
      } else {
        Object.keys(obj).forEach(key => {
          cb(obj[key], key);
        });
      }
    }
  },

  some: function some(obj, cb) {
    if (obj) {
      if (Array.isArray(obj)) {
        return obj.some(cb);
      }
      return Object.keys(obj).some(key => cb(obj[key], key));
    }
    return false;
  },

  every: function every(obj, cb) {
    if (obj) {
      if (Array.isArray(obj)) {
        return obj.every(cb);
      }
      return Object.keys(obj).every(key => cb(obj[key], key));
    }
    return true;
  },

  map: function map(obj, cb) {
    if (obj) {
      if (Array.isArray(obj)) {
        return obj.map(cb);
      }
      return Object.keys(obj).map(key => cb(obj[key], key));
    }
    return [];
  },

  keyBy(a, p) {
    return a.reduce((o, v) => {
      o[v[p]] = v;
      return o;
    }, {});
  },

  isEqual: function isEqual(a, b) {
    const aType = typeof a;
    const bType = typeof b;
    const aArray = Array.isArray(a);
    const bArray = Array.isArray(b);
    let keys;

    if (aType !== bType) {
      return false;
    }
    switch (typeof a) {
      case 'object':
        if (aArray || bArray) {
          if (aArray && bArray) {
            return (
              a.length === b.length &&
              a.every((aValue, index) => {
                const bValue = b[index];
                return _.isEqual(aValue, bValue);
              })
            );
          }
          return false;
        }

        if (a === null || b === null) {
          return a === b;
        }

        // Compare object keys and values
        keys = Object.keys(a);

        if (Object.keys(b).length !== keys.length) {
          return false;
        }

        for (const key of keys) {
          if (!b.hasOwnProperty(key)) {
            return false;
          }
        }

        return _.every(a, (aValue, key) => {
          const bValue = b[key];
          return _.isEqual(aValue, bValue);
        });

      default:
        return a === b;
    }
  },

  escapeHtml(html) {
    const regexResult = escapeHtmlRegex.exec(html);
    if (!regexResult) return html;

    let result = '';
    let escape = '';
    let lastIndex = 0;
    let i = regexResult.index;
    for (; i < html.length; i++) {
      switch (html.charAt(i)) {
        case '"':
          escape = '&quot;';
          break;
        case '&':
          escape = '&amp;';
          break;
        case '\'':
          escape = '&apos;';
          break;
        case '<':
          escape = '&lt;';
          break;
        case '>':
          escape = '&gt;';
          break;
        default:
          continue;
      }
      if (lastIndex !== i) result += html.substring(lastIndex, i);
      lastIndex = i + 1;
      result += escape;
    }
    if (lastIndex !== i) return result + html.substring(lastIndex, i);
    return result;
  },

  strcmp(a, b) {
    if (a < b) return -1;
    if (a > b) return 1;
    return 0;
  },

  isUndefined(val) {
    return toString.call(val) === '[object Undefined]';
  },

  isObject(val) {
    return toString.call(val) === '[object Object]';
  },

  deepMerge() {
    const target = arguments[0] || {};
    const {length} = arguments;
    // eslint-disable-next-line one-var
    let src, clone, copyIsArray;

    function assignValue(val, key) {
      // [exceljs-hardened] CVE fix — prototype pollution
      if (UNSAFE_MERGE_KEYS.has(key)) {
        return;
      }
      src = target[key];
      copyIsArray = Array.isArray(val);
      if (_.isObject(val) || copyIsArray) {
        if (copyIsArray) {
          copyIsArray = false;
          clone = src && Array.isArray(src) ? src : [];
        } else {
          clone = src && _.isObject(src) ? src : {};
        }
        target[key] = _.deepMerge(clone, val);
      } else if (!_.isUndefined(val)) {
        target[key] = val;
      }
    }

    for (let i = 0; i < length; i++) {
      _.each(arguments[i], assignValue);
    }
    return target;
  },
};

module.exports = _;
