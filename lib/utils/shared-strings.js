'use strict';

var SharedStrings = module.exports = function() {
  this._values = [];
  this._totalRefs = 0;
  this._hash = {};
};

SharedStrings.prototype = {
  get count() {
    return this._values.length;
  },
  get values() {
    return this._values;
  },
  get totalRefs() {
    return this._totalRefs;
  },

  getString: function(index) {
    return this._values[index];
  },

  add: function(value) {
    var index = this._hash[value];
    if (index === undefined) {
      index = this._hash[value] = this._values.length;
      this._values.push(value);
    }
    this._totalRefs++;
    return index;
  }
};
