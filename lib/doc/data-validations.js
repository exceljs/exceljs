/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

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
