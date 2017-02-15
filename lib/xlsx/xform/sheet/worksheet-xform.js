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
var _ = require('../../../utils/under-dash');

var utils = require('../../../utils/utils');
var XmlStream = require('../../../utils/xml-stream');

var Merges = require('./merges');
var HyperlinkMap = require('./hyperlink-map');

var BaseXform = require('../base-xform');
var ListXform = require('../list-xform');
var RowXform = require('./row-xform');
var ColXform = require('./col-xform');
var DimensionXform = require('./dimension-xform');
var HyperlinkXform = require('./hyperlink-xform');
var MergeCellXform = require('./merge-cell-xform');
var DataValidationsXform = require('./data-validations-xform');
var SheetPropertiesXform = require('./sheet-properties-xform');
var SheetFormatPropertiesXform = require('./sheet-format-properties-xform');
var SheetViewXform = require('./sheet-view-xform');
var PageMarginsXform = require('./page-margins-xform');
var PageSetupXform = require('./page-setup-xform');
var PrintOptionsXform = require('./print-options-xform');

var WorkSheetXform = module.exports = function() {
  this.map = {
    sheetPr: new SheetPropertiesXform(),
    dimension: new DimensionXform(),
    sheetViews: new ListXform({tag: 'sheetViews', count: false, childXform: new SheetViewXform()}),
    sheetFormatPr: new SheetFormatPropertiesXform(),
    cols: new ListXform({tag: 'cols', count: false, childXform: new ColXform()}),
    sheetData: new ListXform({tag: 'sheetData', count: false, empty: true, childXform: new RowXform()}),
    mergeCells: new ListXform({tag: 'mergeCells', count: true, childXform: new MergeCellXform()}),
    hyperlinks: new ListXform({tag: 'hyperlinks', count: false, childXform: new HyperlinkXform()}),
    pageMargins: new PageMarginsXform(),
    dataValidations: new DataValidationsXform(),
    pageSetup: new PageSetupXform(),
    printOptions: new PrintOptionsXform()
  };
};

utils.inherits(WorkSheetXform, BaseXform, {
  WORKSHEET_ATTRIBUTES: {
    'xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'mc:Ignorable': 'x14ac',
    'xmlns:x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'
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

    var sheetFormatPropertiesModel = model.properties ? {
      defaultRowHeight: model.properties.defaultRowHeight,
      dyDescent: model.properties.dyDescent,
      outlineLevelCol: model.properties.outlineLevelCol,
      outlineLevelRow: model.properties.outlineLevelRow
    } : undefined;
    var sheetPropertiesModel = {
      tabColor: model.properties && model.properties.tabColor,
      pageSetup: model.pageSetup && model.pageSetup.fitToPage ? {
        fitToPage: model.pageSetup.fitToPage
      } : undefined
    };
    var pageMarginsModel = model.pageSetup && model.pageSetup.margins;
    var printOptionsModel = {
      showRowColHeaders: model.showRowColHeaders,
      showGridLines: model.showGridLines,
      horizontalCentered: model.horizontalCentered,
      verticalCentered: model.verticalCentered
    };

    this.map.sheetPr.render(xmlStream, sheetPropertiesModel);
    this.map.dimension.render(xmlStream, model.dimensions);
    this.map.sheetViews.render(xmlStream, model.views);
    this.map.sheetFormatPr.render(xmlStream, sheetFormatPropertiesModel);
    this.map.cols.render(xmlStream, model.cols);
    this.map.sheetData.render(xmlStream, model.rows);
    this.map.mergeCells.render(xmlStream, model.mergeCells);
    this.map.dataValidations.render(xmlStream, model.dataValidations);
    this.map.hyperlinks.render(xmlStream, model.hyperlinks);//For some reason hyperlinks have to be after the data validations
    this.map.pageMargins.render(xmlStream, pageMarginsModel);
    this.map.printOptions.render(xmlStream, printOptionsModel);
    this.map.pageSetup.render(xmlStream, model.pageSetup);

    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else if (node.name === 'worksheet') {
      _.each(this.map, function(xform) {
        xform.reset();
      });
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
          var properties = this.map.sheetFormatPr.model;
          if (this.map.sheetPr.model && this.map.sheetPr.model.tabColor) {
            properties.tabColor = this.map.sheetPr.model.tabColor;
          }
          var sheetProperties = {
            fitToPage: (this.map.sheetPr.model && this.map.sheetPr.model.pageSetup && this.map.sheetPr.model.pageSetup.fitToPage) || false,
            margins: this.map.pageMargins.model
          };
          var pageSetup = Object.assign(
            sheetProperties,
            this.map.pageSetup.model,
            this.map.printOptions.model
          );
          this.model = {
            dimensions: this.map.dimension.model,
            cols: this.map.cols.model,
            rows: this.map.sheetData.model,
            mergeCells: this.map.mergeCells.model,
            hyperlinks: this.map.hyperlinks.model,
            dataValidations: this.map.dataValidations.model,
            properties: properties,
            views: this.map.sheetViews.model,
            pageSetup: pageSetup
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
    model.rows.forEach(function (row) {
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

