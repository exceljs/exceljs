/**
 * Copyright (c) 2016 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 */
'use strict';

var _ = require('../../../utils/under-dash');

var utils = require('../../../utils/utils');
var colCache = require('../../../utils/col-cache');
var XmlStream = require('../../../utils/xml-stream');

var BaseXform = require('../base-xform');
var StaticXform = require('../static-xform');
var ListXform = require('../list-xform');
var DefinedNameXform = require('./defined-name-xform');
var SheetXform = require('./sheet-xform');
var WorkbookViewXform = require('./workbook-view-xform');
var WorkbookPropertiesXform = require('./workbook-properties-xform');

var WorkbookXform = module.exports = function() {
  this.map = {
    fileVersion: WorkbookXform.STATIC_XFORMS.fileVersion,
    workbookPr: new WorkbookPropertiesXform(),
    bookViews: new ListXform({tag: 'bookViews', count: false, childXform: new WorkbookViewXform()}),
    sheets: new ListXform({tag: 'sheets', count: false, childXform: new SheetXform()}),
    definedNames: new ListXform({tag: 'definedNames', count: false, childXform: new DefinedNameXform()}),
    calcPr: WorkbookXform.STATIC_XFORMS.calcPr
  };
};

utils.inherits(WorkbookXform, BaseXform, {
  WORKBOOK_ATTRIBUTES: {
    'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'mc:Ignorable': 'x15',
    'xmlns:x15': 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'
  },
  STATIC_XFORMS: {
    fileVersion: new StaticXform({tag: 'fileVersion', $: {appName: 'xl', lastEdited: 5, lowestEdited: 5, rupBuild: 9303}}),
    calcPr: new StaticXform({tag: 'calcPr', $: {calcId: 171027}})
  }
},{

  prepare: function(model) {
    model.sheets = model.worksheets;

    // collate all the print areas from all of the sheets and add them to the defined names
    var printAreas = [];
    var index = 0; // sheets is sparse array - calc index manually
    model.sheets.forEach(function(sheet) {
      if (sheet.pageSetup && sheet.pageSetup.printArea) {
        var definedName = {
          name: '_xlnm.Print_Area',
          ranges: ['\'' + sheet.name + '\'!' + sheet.pageSetup.printArea],
          localSheetId: index
        };
        printAreas.push(definedName);
      }
      index++;
    });
    if (printAreas.length) {
      model.definedNames = model.definedNames.concat(printAreas);
    }
  },

  render: function(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode('workbook', WorkbookXform.WORKBOOK_ATTRIBUTES);

    this.map.fileVersion.render(xmlStream);
    this.map.workbookPr.render(xmlStream, model.properties);
    this.map.bookViews.render(xmlStream, model.views);
    this.map.sheets.render(xmlStream, model.sheets);
    this.map.definedNames.render(xmlStream, model.definedNames);
    this.map.calcPr.render(xmlStream);

    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      switch(node.name) {
        case 'workbook':
          return true;
        default:
          this.parser = this.map[node.name];
          if (this.parser) {
            this.parser.parseOpen(node);
          }
          return true;
      }
    }
  },
  parseText: function(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },
  parseClose: function(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    } else {
      switch(name) {
        case 'workbook':
          this.model = {
            sheets: this.map.sheets.model,
            properties: this.map.workbookPr.model,
            views: this.map.bookViews.model
          };
          if (this.map.definedNames.model) {
            this.model.definedNames = this.map.definedNames.model;
          }

          return false;
        default:
          // not quite sure how we get here!
          return true;
      }
    }
  },

  reconcile: function(model) {
    var rels = model.workbookRels.reduce(function(map, rel) {
      map[rel.rId] = rel;
      return map;
    }, {});

    // reconcile sheet ids, rIds and names
    var worksheets = [];
    var worksheet;
    var index = 0;

    model.sheets.forEach(function(sheet) {
      var rel = rels[sheet.rId];
      if (!rel) {
        return;
      }
      worksheet = model.worksheetHash['xl/' + rel.target];
      worksheet.name = sheet.name;
      worksheet.id = sheet.id;
      worksheets[index++] = worksheet;
    });

    // reconcile print areas
    var definedNames = [];
    _.each(model.definedNames, function(definedName) {
      if (definedName.name === '_xlnm.Print_Area') {
        worksheet = worksheets[definedName.localSheetId];
        if (worksheet) {
          if (!worksheet.pageSetup) { worksheet.pageSetup = {}; }
          var range  =  colCache.decodeEx(definedName.ranges[0]);
          worksheet.pageSetup.printArea = range.dimensions;
        }
      } else {
        definedNames.push(definedName);
      }
    });
    model.definedNames = definedNames;
  }
});
