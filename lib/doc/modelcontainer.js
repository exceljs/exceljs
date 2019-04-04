'use strict';

const XLSX = require('./../xlsx/xlsx');

const ModelContainer = (module.exports = function(model) {
  this.model = model;
});

ModelContainer.prototype = {
  get xlsx() {
    return this._xlsx || (this._xlsx = new XLSX(this));
  },
};
