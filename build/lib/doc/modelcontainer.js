/**
 * Copyright (c) 2014-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var XLSX = require('./../xlsx/xlsx');

var ModelContainer = module.exports = function (model) {
  this.model = model;
};

ModelContainer.prototype = {
  get xlsx() {
    return this._xlsx || (this._xlsx = new XLSX(this));
  }
};
//# sourceMappingURL=modelcontainer.js.map
