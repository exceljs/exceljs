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
        return _.every(a, (aValue, key) => {
          const bValue = b[key];
          return _.isEqual(aValue, bValue);
        });

      default:
        return a === b;
    }
  },

  escapeHtml(html) {
    return html
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  },
  cloneDeep: function cloneDeep(obj, preserveUndefined) {
    if (preserveUndefined === undefined) {
      preserveUndefined = true;
    }
    let clone;
    if (obj === null) {
      return null;
    }
    if (obj instanceof Date) {
      return obj;
    }
    if (obj instanceof Array) {
      clone = [];
    } else if (typeof obj === 'object') {
      clone = {};
    } else {
      return obj;
    }
   _.each(obj, (value, name) => {
            if (value !== undefined) {
                clone[name] = cloneDeep(value, preserveUndefined);
            } else if (preserveUndefined) {
                clone[name] = undefined;
            }
        });
    return clone;
  },
  strcmp(a, b) {
    if (a < b) return -1;
    if (a > b) return 1;
    return 0;
  },
};

module.exports = _;
