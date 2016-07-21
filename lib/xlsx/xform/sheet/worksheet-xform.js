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

var events = require('events');
var _ = require('underscore');

var utils = require('../../../utils/utils');
var XmlStream = require('../../../utils/xml-stream');

var Merges = require('./merges');
var HyperlinkMap = require('./hyperlink-map');

var BaseXform = require('../base-xform');
var StaticXform = require('../static-xform');
var ListXform = require('../list-xform');
var RowXform = require('./row-xform');
var ColXform = require('./col-xform');
var DimensionXform = require('./dimension-xform');
var HyperlinkXform = require('./hyperlink-xform');
var MergeCellXform = require('./merge-cell-xform');
var DataValidationsXform = require('./data-validations-xform');
var SheetPropertiesXform = require('./sheet-properties-xform');
var SheetFormatPropertiesXform = require('./sheet-format-properties-xform');

var WorkSheetXform = module.exports = function() {

  this.map = {
    sheetPr: new SheetPropertiesXform(),
    dimension: new DimensionXform(),
    sheetViews: WorkSheetXform.STATIC_XFORMS.sheetViews,
    sheetFormatPr: new SheetFormatPropertiesXform(),
    cols: new ListXform({tag: 'cols', count: false, childXform: new ColXform()}),
    sheetData: new ListXform({tag: 'sheetData', count: false, childXform: new RowXform()}),
    mergeCells: new ListXform({tag: 'mergeCells', count: true, childXform: new MergeCellXform()}),
    hyperlinks: new ListXform({tag: 'hyperlinks', count: false, childXform: new HyperlinkXform()}),
    pageMargins: WorkSheetXform.STATIC_XFORMS.pageMargins,
    dataValidations: new DataValidationsXform()
  };
  
};

utils.inherits(WorkSheetXform, BaseXform, {
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
    model.hyperlinks = options.hyperlinks = [];

    this.map.cols.prepare(model.cols, options);
    this.map.sheetData.prepare(model.rows, options);

    model.mergeCells = options.merges.mergeCells;
  },

  render: function(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode('worksheet', WorkSheetXform.WORKSHEET_ATTRIBUTES);

    this.map.sheetPr.render(xmlStream, model.properties);
    this.map.dimension.render(xmlStream, model.dimensions);
    this.map.sheetViews.render(xmlStream);
    this.map.sheetFormatPr.render(xmlStream, model.properties);
    this.map.cols.render(xmlStream, model.cols);
    this.map.sheetData.render(xmlStream, model.rows);
    this.map.mergeCells.render(xmlStream, model.mergeCells);
    this.map.hyperlinks.render(xmlStream, model.hyperlinks);
    this.map.dataValidations.render(xmlStream, model.dataValidations);
    this.map.pageMargins.render(xmlStream);

    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else if (node.name === 'worksheet') {
      return true;
    } else {
      this.parser = this.map[node.name];
      if (this.parser) {
        this.parser.parseOpen(node);
      }
      return true;
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
          this.model = {
            dimensions: this.map.dimension.model,
            cols: this.map.cols.model,
            rows: this.map.sheetData.model,
            mergeCells: this.map.mergeCells.model,
            hyperlinks: this.map.hyperlinks.model,
            dataValidations: this.map.dataValidations.model,
            properties: _.extend({}, this.map.sheetPr.model, this.map.sheetFormatPr.model)
          };
          return false;
        default:
          // not quite sure how we get here!
          return true;
      }
    }
  },

  reconcile: function(model, options) {
    // options.merges = new Merges();
    // options.merges.reconcile(model.mergeCells, model.rows);
    options.hyperlinkMap  = new HyperlinkMap(model.relationships, model.hyperlinks);

    // compact the rows and cells
    model.rows = model.rows && model.rows.filter(function (row) {
      return row;
    }) || [];
    _.each(model.rows, function (row) {
      row.cells = row.cells && row.cells.filter(function (cell) {
        return cell;
      }) || [];
    });
    
    this.map.cols.reconcile(model.cols, options);
    this.map.sheetData.reconcile(model.rows, options);

    model.relationships = undefined;
    model.hyperlinks = undefined;
  }
});

