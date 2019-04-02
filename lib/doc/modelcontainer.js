'use strict';

var XLSX = require('./../xlsx/xlsx');

var ModelContainer = module.exports = function(model) {
  this.model = model;
};


ModelContainer.prototype = {
  get xlsx() {
    return this._xlsx || (this._xlsx = new XLSX(this));
  }
};
