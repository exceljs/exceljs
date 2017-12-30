/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var DataValidations = module.exports = function (model) {
  this.model = model || {};
};

DataValidations.prototype = {
  add: function add(address, validation) {
    return this.model[address] = validation;
  },
  find: function find(address) {
    return this.model[address];
  },
  remove: function remove(address) {
    this.model[address] = undefined;
  }
};
//# sourceMappingURL=data-validations.js.map
