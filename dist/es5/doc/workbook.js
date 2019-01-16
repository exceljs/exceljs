/**
 * Copyright (c) 2014-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var Worksheet = require('./worksheet');
var DefinedNames = require('./defined-names');
var XLSX = require('./../xlsx/xlsx');
var CSV = require('./../csv/csv');

// Workbook requirements
//  Load and Save from file and stream
//  Access/Add/Delete individual worksheets
//  Manage String table, Hyperlink table, etc.
//  Manage scaffolding for contained objects to write to/read from

var Workbook = module.exports = function () {
  this.created = new Date();
  this.modified = this.created;
  this.properties = {};
  this._worksheets = [];
  this.views = [];
  this.media = [];
  this._definedNames = new DefinedNames();
};

Workbook.prototype = {
  get xlsx() {
    if (!this._xlsx) this._xlsx = new XLSX(this);
    return this._xlsx;
  },
  get csv() {
    if (!this._csv) this._csv = new CSV(this);
    return this._csv;
  },
  get nextId() {
    // find the next unique spot to add worksheet
    var i;
    for (i = 1; i < this._worksheets.length; i++) {
      if (!this._worksheets[i]) {
        return i;
      }
    }
    return this._worksheets.length || 1;
  },
  addWorksheet: function addWorksheet(name, options) {
    var id = this.nextId;
    name = name || 'sheet' + id;

    // if options is a color, call it tabColor (and signal deprecated message)
    if (options) {
      if (typeof options === 'string') {
        // eslint-disable-next-line no-console
        console.trace('tabColor argument is now deprecated. Please use workbook.addWorksheet(name, {properties: { tabColor: { argb: "rbg value" } }');
        options = {
          properties: {
            tabColor: { argb: options }
          }
        };
      } else if (options.argb || options.theme || options.indexed) {
        // eslint-disable-next-line no-console
        console.trace('tabColor argument is now deprecated. Please use workbook.addWorksheet(name, {properties: { tabColor: { ... } }');
        options = {
          properties: {
            tabColor: options
          }
        };
      }
    }

    var lastOrderNo = this._worksheets.reduce(function (acc, ws) {
      return (ws && ws.orderNo) > acc ? ws.orderNo : acc;
    }, 0);
    var worksheetOptions = Object.assign({}, options, {
      id: id,
      name: name,
      orderNo: lastOrderNo + 1,
      workbook: this
    });

    var worksheet = new Worksheet(worksheetOptions);

    this._worksheets[id] = worksheet;
    return worksheet;
  },
  removeWorksheetEx: function removeWorksheetEx(worksheet) {
    delete this._worksheets[worksheet.id];
  },
  removeWorksheet: function removeWorksheet(id) {
    var worksheet = this.getWorksheet(id);
    if (worksheet) {
      worksheet.destroy();
    }
  },

  getWorksheet: function getWorksheet(id) {
    if (id === undefined) {
      return this._worksheets.find(function (worksheet) {
        return worksheet;
      });
    } else if (typeof id === 'number') {
      return this._worksheets[id];
    } else if (typeof id === 'string') {
      return this._worksheets.find(function (worksheet) {
        return worksheet && worksheet.name === id;
      });
    }
    return undefined;
  },

  get worksheets() {
    // return a clone of _worksheets
    return this._worksheets.slice(1).sort(function (a, b) {
      return a.orderNo - b.orderNo;
    }).filter(Boolean);
  },

  eachSheet: function eachSheet(iteratee) {
    this.worksheets.forEach(function (sheet) {
      iteratee(sheet, sheet.id);
    });
  },

  get definedNames() {
    return this._definedNames;
  },

  clearThemes: function clearThemes() {
    // Note: themes are not an exposed feature, meddle at your peril!
    this._themes = undefined;
  },

  addImage: function addImage(image) {
    // TODO:  validation?
    var id = this.media.length;
    this.media.push(Object.assign({}, image, { type: 'image' }));
    return id;
  },

  getImage: function getImage(id) {
    return this.media[id];
  },


  get model() {
    return {
      creator: this.creator || 'Unknown',
      lastModifiedBy: this.lastModifiedBy || 'Unknown',
      lastPrinted: this.lastPrinted,
      created: this.created,
      modified: this.modified,
      properties: this.properties,
      worksheets: this.worksheets.map(function (worksheet) {
        return worksheet.model;
      }),
      sheets: this.worksheets.map(function (ws) {
        return ws.model;
      }).filter(Boolean),
      definedNames: this._definedNames.model,
      views: this.views,
      company: this.company,
      manager: this.manager,
      title: this.title,
      subject: this.subject,
      keywords: this.keywords,
      category: this.category,
      description: this.description,
      language: this.language,
      revision: this.revision,
      contentStatus: this.contentStatus,
      themes: this._themes,
      media: this.media
    };
  },
  set model(value) {
    var _this = this;

    this.creator = value.creator;
    this.lastModifiedBy = value.lastModifiedBy;
    this.lastPrinted = value.lastPrinted;
    this.created = value.created;
    this.modified = value.modified;
    this.company = value.company;
    this.manager = value.manager;
    this.title = value.title;
    this.subject = value.subject;
    this.keywords = value.keywords;
    this.category = value.category;
    this.description = value.description;
    this.language = value.language;
    this.revision = value.revision;
    this.contentStatus = value.contentStatus;

    this.properties = value.properties;
    this._worksheets = [];
    value.worksheets.forEach(function (worksheetModel) {
      var id = worksheetModel.id;
      var name = worksheetModel.name;
      var orderNo = value.sheets.findIndex(function (ws) {
        return ws.id === id;
      });
      var worksheet = _this._worksheets[id] = new Worksheet({
        id: id,
        name: name,
        orderNo: orderNo,
        workbook: _this
      });

      worksheet.model = worksheetModel;
    });

    this._definedNames.model = value.definedNames;
    this.views = value.views;
    this._themes = value.themes;
    this.media = value.media || [];
  }
};
//# sourceMappingURL=workbook.js.map
