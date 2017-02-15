/**
 * Copyright (c) 2015 Guyon Roche
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
 * THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 */
'use strict';

var fs = require('fs');
var Archiver = require('archiver');

var utils = require('../../utils/utils');
var StreamBuf = require('../../utils/stream-buf');

var RelType = require('../../xlsx/rel-type');
var StylesXform = require('../../xlsx/xform/style/styles-xform');
var SharedStrings = require('../../utils/shared-strings');
var DefinedNames = require('../../doc/defined-names');

var CoreXform = require('../../xlsx/xform/core/core-xform');
var RelationshipsXform = require('../../xlsx/xform/core/relationships-xform');
var ContentTypesXform = require('../../xlsx/xform/core/content-types-xform');
var AppXform = require('../../xlsx/xform/core/app-xform');
var WorkbookXform = require('../../xlsx/xform/book/workbook-xform');
var SharedStringsXform = require('../../xlsx/xform/strings/shared-strings-xform');

var WorksheetWriter = require('./worksheet-writer');

var theme1Xml = require('../../xlsx/xml/theme1.xml.js');

var WorkbookWriter = module.exports = function(options) {
  options = options || {};
  //console.log(JSON.stringify(options, null, '    '))

  this.created = options.created || new Date();
  this.modified = options.modified || this.created;
  this.creator = options.creator || 'ExcelJS';
  this.lastModifiedBy = options.lastModifiedBy || 'ExcelJS';
  this.lastPrinted = options.lastPrinted;

  // using shared strings creates a smaller xlsx file but may use more memory
  this.useSharedStrings = options.useSharedStrings || false;
  this.sharedStrings = new SharedStrings();

  // style manager
  this.styles = options.useStyles ? new StylesXform(true) : new StylesXform.Mock(true);

  // defined names
  this._definedNames = new DefinedNames();

  this._worksheets = [];
  this.views = [];

  this.zip = Archiver('zip');
  if (options.stream) {
    this.stream = options.stream;
  } else if (options.filename) {
    this.stream = fs.createWriteStream(options.filename);
  } else {
    this.stream = new StreamBuf();
  }
  this.zip.pipe(this.stream);

  // these bits can be added right now
  this.promise = Promise.all([
    this.addThemes(),
    this.addOfficeRels()
  ]);

};
WorkbookWriter.prototype = {
  get definedNames() {
    return this._definedNames;
  },

  _openStream: function(path) {
    var self = this;
    var stream = new StreamBuf({bufSize: 65536, batch: true});
    self.zip.append(stream, { name: path });
    stream.on('finish', function() {
      stream.emit('zipped');
    });
    return stream;
  },
  _commitWorksheets: function() {
    var commitWorksheet = function(worksheet) {
      if (!worksheet.committed) {
        return new Promise(function(resolve) {
          worksheet.stream.on('zipped', function() {
            resolve();
          });
          worksheet.commit();
        });
      } else {
        return Promise.resolve();
      }
    };
    // if there are any uncommitted worksheets, commit them now and wait
    var promises = this._worksheets.map(commitWorksheet);
    if (promises.length) {
      return Promise.all(promises);
    } else {
      return Promise.resolve();
    }
  },
  commit: function() {
    var self = this;

    // commit all worksheets, then add suplimentary files
    return this.promise.then(function() {
        return self._commitWorksheets();
      })
      .then(function() {
        return Promise.all([
          self.addContentTypes(),
          self.addApp(),
          self.addCore(),
          self.addSharedStrings(),
          self.addStyles(),
          self.addWorkbookRels()
        ]);
      })
      .then(function() {
        return self.addWorkbook();
      })
      .then(function(){
        return self._finalize();
      });
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
    // it's possible to add a worksheet with different than default
    // shared string handling
    // in fact, it's even possible to switch it mid-sheet
    options = options || {};
    var useSharedStrings = options.useSharedStrings !== undefined ?
      options.useSharedStrings :
      this.useSharedStrings;

    if (options.tabColor) {
      console.trace('tabColor option has moved to { properties: tabColor: {...} }');
      options.properties  = Object.assign({
        tabColor: options.tabColor
      }, options.properties);
    }

    var id = this.nextId;
    name = name || 'sheet' + id;

    var worksheet = new WorksheetWriter({
      id: id,
      name: name,
      workbook: this,
      useSharedStrings: useSharedStrings,
      properties: options.properties,
      pageSetup: options.pageSetup,
      views: options.views
    });

    this._worksheets[id] = worksheet;
    return worksheet;
  },
  getWorksheet: function(id) {
    if (id === undefined) {
      return this._worksheets.find(function() { return true; });
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
  addStyles: function() {
    var self = this;
    return new Promise(function(resolve) {
      self.zip.append(self.styles.xml, {name: 'xl/styles.xml'});
      resolve();
    });
  },
  addThemes: function() {
    var self = this;
    return new Promise(function(resolve) {
      self.zip.append(theme1Xml, { name: 'xl/theme/theme1.xml' });
      resolve();
    });
  },
  addOfficeRels: function() {
    var self = this;
    return new Promise(function(resolve) {
      var xform = new RelationshipsXform();
      var xml = xform.toXml([
        {rId: 'rId1', type: RelType.OfficeDocument, target: 'xl/workbook.xml'}
      ]);
      self.zip.append(xml, {name: '/_rels/.rels'});
      resolve();
    });
  },

  addContentTypes: function() {
    var self = this;
    return new Promise(function(resolve) {
      var model = {
        worksheets: self._worksheets.filter(Boolean)
      };
      var xform = new ContentTypesXform();
      var xml = xform.toXml(model);
      self.zip.append(xml, {name: '[Content_Types].xml'});
      resolve();
    });
  },
  addApp: function() {
    var self = this;
    return new Promise(function(resolve) {
      var model = {
        worksheets: self._worksheets.filter(Boolean)
      };
      var xform = new AppXform();
      var xml = xform.toXml(model);
      self.zip.append(xml, {name: 'docProps/app.xml'});
      resolve();
    });
  },
  addCore: function() {
    var self = this;
    return new Promise(function(resolve) {
      var coreXform = new CoreXform();
      var xml = coreXform.toXml(self);
      self.zip.append(xml, {name: 'docProps/core.xml'});
      resolve();
    });
  },
  addSharedStrings: function() {
    var self = this;
    if (this.sharedStrings.count) {
      return new Promise(function(resolve) {
        var sharedStringsXform = new SharedStringsXform();
        var xml = sharedStringsXform.toXml(self.sharedStrings);
        self.zip.append(xml, {name: '/xl/sharedStrings.xml'});
        resolve();
      });
    } else {
      return Promise.resolve();
    }
  },
  addWorkbookRels: function() {
    var self = this;
    var count = 1;
    var relationships = [
      {rId: 'rId' + (count++), type: RelType.Styles, target: 'styles.xml'},
      {rId: 'rId' + (count++), type: RelType.Theme, target: 'theme/theme1.xml'}
    ];
    if (this.sharedStrings.count) {
      relationships.push(
        {rId: 'rId' + (count++), type: RelType.SharedStrings, target: 'sharedStrings.xml'}
      );
    }
    this._worksheets.forEach(function (worksheet) {
      if (worksheet) {
        worksheet.rId = 'rId' + (count++);
        relationships.push(
          {rId: worksheet.rId, type: RelType.Worksheet, target: 'worksheets/sheet' + worksheet.id + '.xml'}
        );
      }
    });
    return new Promise(function(resolve) {
      var xform = new RelationshipsXform();
      var xml = xform.toXml(relationships);
      self.zip.append(xml, {name: '/xl/_rels/workbook.xml.rels'});
      resolve();
    });
  },
  addWorkbook: function() {
    var zip = this.zip;
    var model = {
      worksheets: this._worksheets.filter(Boolean),
      definedNames: this._definedNames.model,
      views: this.views,
      properties: {}
    };

    return new Promise(function(resolve) {
      var xform = new WorkbookXform();
      xform.prepare(model);
      zip.append(xform.toXml(model), {name: '/xl/workbook.xml'});
      resolve();
    });
  },
  _finalize: function() {
    var self = this;
    return new Promise(function(resolve, reject) {
      self.stream.on('error', function(error){
        reject(error);
      });
      self.zip.on('finish', function(){
        resolve(self);
      });
      self.zip.on('error', function(error){
        reject(error);
      });

      self.zip.finalize();
    });
  }
};
