'use strict';

const DataValidations = (module.exports = function(model) {
  this.model = model || {};
});

DataValidations.prototype = {
  add(address, validation) {
    return (this.model[address] = validation);
  },
  find(address) {
    return this.model[address];
  },
  remove(address) {
    this.model[address] = undefined;
  },
};
