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

var _ = require('underscore');

var utils = require('../../../utils/utils');
var XmlStream = require('../../../utils/xml-stream');

var BaseXform = require('../base-xform');
var StaticXform = require('../static-xform');
var ListXform = require('../list-xform');
var DefinedNameXform = require('./defined-name-xform');
var SheetXform = require('./sheet-xform');

// <workbook
// xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
// xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
//   <fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9303"/>
//   <workbookPr defaultThemeVersion="124226"/>
//   <bookViews>
//   <workbookView xWindow="480" yWindow="60" windowWidth="18195" windowHeight="8505"/>
//   </bookViews>
//   <sheets>
//     <% _.each(worksheets, function(worksheet) { %><sheet name="<%=worksheet.name%>" sheetId="<%=worksheet.id%>" r:id="<%=worksheet.rId%>"/>
//     <% }); %>
//   </sheets>
//  <%=definedNames.xml%>
//   <calcPr calcId="145621"/>
// </workbook>

// This comes in some xml, not sure what for yet
// <extLst>
//   <ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
//     <x15:workbookPr chartTrackingRefBase="1"/>
//   </ext>
// </extLst>



var WorkbookXform = module.exports = function() {
  this.map = {
    fileVersion: WorkbookXform.STATIC_XFORMS.fileVersion,
    workbookPr: WorkbookXform.STATIC_XFORMS.workbookPr,
    bookViews: WorkbookXform.STATIC_XFORMS.bookViews,
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
    workbookPr: new StaticXform({tag: 'workbookPr', $: {defaultThemeVersion: 164011, filterPrivacy: 1}}),
    bookViews: new StaticXform({tag: 'bookViews', c: [{tag: 'workbookView',  $: {xWindow:0, yWindow: 0, windowWidth: 22260, windowHeight: 12648}}]}),
    calcPr: new StaticXform({tag: 'calcPr', $: {calcId: 171027}})
  }
},{

  prepare: function(model) {
    model.sheets = model.worksheets;
  },

  render: function(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode('workbook', WorkbookXform.WORKBOOK_ATTRIBUTES);

    this.map.fileVersion.render(xmlStream);
    this.map.workbookPr.render(xmlStream);
    this.map.bookViews.render(xmlStream);
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
            sheets: this.map.sheets.model
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
    _.each(model.sheets, function (sheet) {
      var rel = rels[sheet.rId];
      if (!rel) {
        return;
      }
      var worksheet = model.worksheetHash['xl/' + rel.target];
      worksheet.name = sheet.name;
      worksheet.id = sheet.id;
    });
    
  }
});
