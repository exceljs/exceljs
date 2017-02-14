/**
 * Copyright (c) 2014 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the 'Software'), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
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

var Workbook = module.exports = function() {
  this.created = new Date();
  this.modified = this.created;
  this.properties = {};
  this._worksheets = [];
  this.views = [];
  this._definedNames = new DefinedNames();
};

Workbook.prototype = {
  get xlsx() {
    return this._xlsx || (this._xlsx = new XLSX(this));
  },
  get csv() {
    return this._csv || (this._csv = new CSV(this));
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
  addWorksheet: function(name, options) {

    var id = this.nextId;
    name = name || 'sheet' + id;
    
    // if options is a color, call it tabColor (and signal deprecated message)
    if (options) {
      if (typeof options === 'string') {
        console.trace('tabColor argument is now deprecated. Please use workbook.addWorksheet(name, {properties: { tabColor: { argb: "rbg value" } }');
        options = {
          properties: {
            tabColor: {argb: options}
          }
        };
      } else if (options.argb || options.theme || options.indexed) {
        console.trace('tabColor argument is now deprecated. Please use workbook.addWorksheet(name, {properties: { tabColor: { ... } }');
        options = {
          properties: {
            tabColor: options
          }
        };
      }
    }

    var worksheetOptions = Object.assign({}, options, {
      id: id,
      name: name,
      workbook: this
    });

    var worksheet = new Worksheet(worksheetOptions);

    this._worksheets[id] = worksheet;
    return worksheet;
  },
  _removeWorksheet: function(worksheet) {
    delete this._worksheets[worksheet.id];
  },
  removeWorksheet: function(id) {
    var worksheet = this.getWorksheet(id);
    if (worksheet) {
      worksheet.destroy();
    }
  },

  getWorksheet: function(id) {
    if (id === undefined) {
      return this._worksheets.find(function(worksheet) { return worksheet; });
    } else if (typeof(id) === 'number') {
      return this._worksheets[id];
    } else if (typeof id === 'string') {
      return this._worksheets.find(function(worksheet) {
        return worksheet && worksheet.name == id;
      });
    } else {
      return undefined;
    }
  },

  get worksheets() {
    // return a clone of _worksheets
    return this._worksheets.filter(Boolean);
  },

  eachSheet: function(iteratee) {
    this._worksheets.forEach(function(sheet) {
      iteratee(sheet, sheet.id);
    });
  },

  get definedNames() {
    return this._definedNames;
  },

  get model() {
    return {
      creator: this.creator || 'Unknown',
      lastModifiedBy: this.lastModifiedBy || 'Unknown',
      lastPrinted: this.lastPrinted,
      created: this.created,
      modified: this.modified,
      properties: this.properties,
      worksheets: this._worksheets.filter(Boolean).map(function(worksheet) { return worksheet.model; }),
      definedNames: this._definedNames.model,
      views: this.views,
      company: this.company,
      manager: this.manager,
      title: this.title,
      subject: this.subject,
      keywords: this.keywords,
      category:  this.category,
      description: this.description,
      language: this.language,
      revision: this.revision,
    };
  },
  set model(value) {
    var self = this;

    // oh-for-a-rest-spread-operator!
    this.creator = value.creator;
    this.lastModifiedBy = value.lastModifiedBy;
    this.lastPrinted = value.lastPrinted;
    this.created = value.created;
    this.modified = value.modified;
    this.company = value.company;
    this.manager  = value.manager;
    this.title  = value.title;
    this.subject  = value.subject;
    this.keywords  = value.keywords;
    this.category  = value.category;
    this.description  = value.description;
    this.language = value.language;
    this.revision = value.revision;

    this.properties = value.properties;
    this._worksheets = [];
    value.worksheets.forEach(function(worksheetModel) {
      var id = worksheetModel.id;
      var name = worksheetModel.name;
      var worksheet = self._worksheets[id] = new Worksheet({
        id: id,
        name: name,
        workbook: self
      });

      worksheet.model = worksheetModel;
    });

    this._definedNames.model = value.definedNames;
    this.views = value.views;
  }
};
