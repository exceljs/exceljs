'use strict';

var verquire = require('./verquire');

var _ = Object.assign({
  get: function get(obj, path, dflt) {
    if (typeof path === 'string') {
      path = path.split('.');
    }
    while (obj && path.length) {
      obj = obj[path.shift()];
    }
    return obj !== undefined ? obj : dflt;
  },

  has: function has(obj, path) {
    var dummy = {};
    return _.get(obj, path, dummy) !== dummy;
  },

  cloneDeep: function cloneDeep(obj, preserveUndefined) {
    if (preserveUndefined === undefined) {
      preserveUndefined = true;
    }
    var clone;
    if (obj instanceof Array) {
      clone = [];
    } else if (obj instanceof Date) {
      return obj;
    } else if (typeof obj === 'object') {
      clone = {};
    } else {
      return obj;
    }
    _.each(obj, function(value, name) {
      if (value !== undefined) {
        clone[name] = cloneDeep(value, preserveUndefined);
      } else if (preserveUndefined) {
        clone[name] = undefined;
      }
    });
    return clone;
  }

}, verquire('utils/under-dash'));

module.exports = _;
