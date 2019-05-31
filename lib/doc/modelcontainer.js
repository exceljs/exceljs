'use strict';

const XLSX = require('./../xlsx/xlsx');

module.exports = class ModelContainer {
  constructor(model) {
    this.model = model;
  }

  get xlsx() {
    if (!this._xlsx) {
      this._xlsx = new XLSX(this);
    }
    return this._xlsx;
  }
};
