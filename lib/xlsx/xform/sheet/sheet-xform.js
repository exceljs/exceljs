/**
 * Copyright (c) 2015 Guyon Roche
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

var events = require('events');

var utils = require('../../../utils/utils');
var XmlStream = require('../../../utils/xml-stream');

var Merges = require('./merges');

var BaseXform = require('../base-xform');
var StaticXform = require('../static-xform');
var ListXform = require('../list-xform');
var RowXform = require('./row-xform');
var ColorXform = require('../style/color-xform');
var ColXform = require('./col-xform');
var DimensionXform = require('./dimension-xform');
var HyperlinkXform = require('./hyperlink-xform');
var MergeCellXform = require('./merge-cell-xform');
var DataValidationsXform = require('./data-validations-xform');

var SheetXform = module.exports = function() {

  this.map = {
    sheetPr: new ListXform({tag: 'sheetPr', count: false, childXform: new ColorXform('tabColor')}),  // necessary?
    dimension: new DimensionXform(),
    sheetViews: SheetXform.STATIC_XFORMS.sheetViews,
    sheetFormatPr: SheetXform.STATIC_XFORMS.sheetFormatPr,
    cols: new ListXform({tag: 'cols', count: false, childXform: new ColXform()}),
    sheetData: new ListXform({tag: 'sheetData', count: false, childXform: new RowXform()}),
    mergeCells: new ListXform({tag: 'MergeCells', count: true, childXform: new MergeCellXform()}),
    hyperlinks: new ListXform({tag: 'hyperlinks', count: false, childXform: new HyperlinkXform()}),
    pageMargins: SheetXform.STATIC_XFORMS.pageMargins,
    dataValidations: new DataValidationsXform()
  };
  
};

utils.inherits(SheetXform, BaseXform, {
  WORKSHEET_ATTRIBUTES: {
    'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'mc:Ignorable': 'x14ac',
    'xmlns:x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'
  },
  STATIC_XFORMS: {
    sheetViews: new StaticXform({tag: 'sheetViews', c: [{tag:'sheetView', $: {workbookViewId:'0'}}]}),
    sheetFormatPr: new StaticXform({tag: 'sheetFormatPr', $: {defaultRowHeight:15, 'x14ac:dyDescent': '0.25'}}),
    pageMargins: new StaticXform({tag: 'pageMargins', $: {left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer:  0.3 }})
  }
},{

  prepare: function(model, options) {
    options.merges = new Merges();
    model.hyperlinks = options.hyperlinks;
    this.map.sheetData.prepare(model.rows, options);
    model.mergeCells = options.merges.mergeCells;
  },

  write: function(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode('worksheet', SheetXform.WORKSHEET_ATTRIBUTES);

    // sheet properties
    if (model.tabColor) {
      this.map.sheetPr.write(xmlStream, [model.tabColor]);
    }
    this.map.dimension.write(xmlStream, model.dimensions);
    this.map.sheetViews.write(xmlStream);
    this.map.sheetFormatPr.write(xmlStream);
    this.map.cols.write(xmlStream, model.cols);
    this.map.sheetData.write(xmlStream, model.rows);
    this.map.mergeCells.write(xmlStream, model.mergeCells);
    this.map.hyperlinks.write(xmlStream, model.hyperlinks);
    this.map.dataValidations.write(xmlStream, model.dataValidations);
    this.map.pageMargins.write(xmlStream);

    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      switch(node.name) {
        case 'worksheet':
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
        case 'worksheet':
          return false;
        default:
          // not quite sure how we get here!
          return true;
      }
    }
  },

  reconcile: function(model, options) {
    options.merges = new Merges(model.mergeCells);
    options.merges.reconcile();

    this.map.sheetData.reconcile(model, options);
  }
});
