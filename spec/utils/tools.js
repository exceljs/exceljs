const MemoryStream = require('memorystream');
const _ = require('./under-dash');

const tools = {
  dtMatcher: /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}[.]\d{3}Z$/,
  fix: function fix(o) {
    // clone the object and replace any date-like strings with new Date()
    let clone;
    if (o instanceof Array) {
      clone = [];
    } else if (typeof o === 'object') {
      clone = {};
    } else if (typeof o === 'string' && tools.dtMatcher.test(o)) {
      return new Date(o);
    } else {
      return o;
    }
    _.each(o, (value, name) => {
      if (value !== undefined) {
        clone[name] = fix(value);
      }
    });
    return clone;
  },

  concatenateFormula() {
    const args = Array.prototype.slice.call(arguments);
    const values = args.map(value => `"${value}"`);
    return {
      formula: `CONCATENATE(${values.join(',')})`,
    };
  },
  cloneByModel(thing1, Type) {
    const {model} = thing1;
    const thing2 = new Type();
    thing2.model = model;
    return Promise.resolve(thing2);
  },
  cloneByStream(thing1, Type, end) {
    return new Promise((resolve, reject) => {
      end = end || 'end';

      const thing2 = new Type();
      const stream = thing2.createInputStream();
      stream.on(end, () => {
        resolve(thing2);
      });
      stream.on('error', error => {
        reject(error);
      });

      const memStream = new MemoryStream();
      memStream.on('error', error => {
        reject(error);
      });
      memStream.pipe(stream);
      thing1.write(memStream).then(() => {
        memStream.end();
      });
    });
  },
  toISODateString(dt) {
    const iso = dt.toISOString();
    const parts = iso.split('T');
    return parts[0];
  },
};

module.exports = tools;
