'use strict';

var DataValidations = module.exports = function(model) {
  this.model = model || {};
};

DataValidations.prototype = {
  add: function(address, validation) {
    return (this.model[address] = validation);
  },
  find: function(address) {
    return this.model[address];
  },
  remove: function(address) {
    this.model[address] = undefined;
  }
};

