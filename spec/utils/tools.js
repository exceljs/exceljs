'use strict';

var _ = require('./under-dash');
var MemoryStream = require('memorystream');

var tools = module.exports = {
  dtMatcher: /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}[.]\d{3}Z$/,
  fix: function fix(o) {
    // clone the object and replace any date-like strings with new Date()
    var clone;
    if (o instanceof Array) {
      clone = [];
    } else if (typeof o === 'object') {
      clone = {};
    } else if ((typeof o === 'string') && tools.dtMatcher.test(o)) {
      return new Date(o);
    } else {
      return o;
    }
    _.each(o, function(value, name) {
      if (value !== undefined) {
        clone[name] = fix(value);
      }
    });
    return clone;
  },

  concatenateFormula: function() {
    var args = Array.prototype.slice.call(arguments);
    var values = args.map(function(value) {
      return '"' + value + '"';
    });
    return {
      formula: 'CONCATENATE(' + values.join(',') + ')'
    };
  },
  cloneByModel: function(thing1, Type) {
    var model = thing1.model;
    var thing2 = new Type();
    thing2.model = model;
    return Promise.resolve(thing2);
  },
  cloneByStream: function(thing1, Type, end) {
    return new Promise(function(resolve, reject) {
    end = end || 'end';

    var thing2 = new Type();
    var stream = thing2.createInputStream();
    stream.on(end, function() {
      resolve(thing2);
    });
    stream.on('error', function(error) {
      reject(error);
    });

    var memStream = new MemoryStream();
    memStream.on('error', function(error) {
      reject(error);
    });
    memStream.pipe(stream);
    thing1.write(memStream)
      .then(function() {
        memStream.end();
      });
    });
  },
  toISODateString: function(dt) {
    var iso = dt.toISOString();
    var parts = iso.split('T');
    return parts[0];
  }
};
