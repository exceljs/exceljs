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

var fs = require('fs');

var utils = require('../../utils/utils');
var colCache = require('../../utils/col-cache');
var Dimensions = require('../../doc/range');

var StringBuf = require('../../utils/string-buf');

var Row = require('../../doc/row');
var Column = require('../../doc/column');

var HyperlinkWriter = require('./hyperlink-writer');
var DataValidations = require('../../doc/data-validations');

var xml = new StringBuf();

// ============================================================================================
// Xforms
var ListXform = require('../../xlsx/xform/list-xform');
var DataValidationsXform = require('../../xlsx/xform/sheet/data-validations-xform');
var SheetPropertiesXform = require('../../xlsx/xform/sheet/sheet-properties-xform');
var SheetFormatPropertiesXform = require('../../xlsx/xform/sheet/sheet-format-properties-xform');
var ColXform = require('../../xlsx/xform/sheet/col-xform');
var RowXform = require('../../xlsx/xform/sheet/row-xform');
var HyperlinkXform = require('../../xlsx/xform/sheet/hyperlink-xform');
var SheetViewXform =  require('../../xlsx/xform/sheet/sheet-view-xform');
var PageMarginsXform =  require('../../xlsx/xform/sheet/page-margins-xform');
var PageSetupXform =  require('../../xlsx/xform/sheet/page-setup-xform');

// since prepare and render is functional, we can use singletons
var xform = {
  dataValidations: new DataValidationsXform(),
  sheetProperties: new SheetPropertiesXform(),
  sheetFormatProperties: new SheetFormatPropertiesXform(),
  columns: new ListXform({tag: 'cols', length: false, childXform: new ColXform()}),
  row: new RowXform(),
  hyperlinks: new ListXform({tag: 'hyperlinks', length: false, childXform: new HyperlinkXform()}),
  sheetViews: new ListXform({tag: 'sheetViews', length: false, childXform: new SheetViewXform()}),
  pageMargins: new PageMarginsXform(),
  pageSeteup: new PageSetupXform()
};


// ============================================================================================

var WorksheetWriter = module.exports = function(options) {
  // in a workbook, each sheet will have a number
  this.id = options.id;

  // and a name
  this.name = options.name || 'Sheet' + this.id;

  // rows are stored here while they need to be worked on.
  // when they are committed, they will be deleted.
  this._rows = [];

  // column definitions
  this._columns = null;

  // column keys (addRow convenience): key ==> this._columns index
  this._keys = {};

  // keep record of all merges
  this._merges = [];
  this._merges.add  = function(){}; // ignore cell instruction

  // keep record of all hyperlinks
  this._hyperlinkWriter = new HyperlinkWriter(options);

  // keep a record of dimensions
  this._dimensions = new Dimensions();

  // first uncommitted row
  this._rowZero = 1;

  // committed flag
  this.committed = false;

  // for data validations
  this.dataValidations = new DataValidations();

  // for default row height, outline levels, etc
  this.properties = Object.assign({}, {
    defaultRowHeight: 15,
    dyDescent: 55,
    outlineLevelCol: 0,
    outlineLevelRow: 0
  }, options.properties);

  // for all things printing
  this.pageSetup = Object.assign({}, {
    margins: {left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer:  0.3 },
    orientation: 'portrait',
    horizontalDpi: 4294967295,
    verticalDpi: 4294967295,
    fitToPage: !!(options.pageSetup && ((options.pageSetup.fitToWidth || options.pageSetup.fitToHeight) && !options.pageSetup.scale)),
    pageOrder: 'downThenOver',
    blackAndWhite: false,
    draft: false,
    cellComments: 'None',
    errors: 'displayed',
    scale: 100,
    fitToWidth: 1,
    fitToHeight: 1,
    paperSize: undefined,
    showRowColHeaders: false,
    showGridLines: false,
    horizontalCentered: false,
    verticalCentered: false,
    rowBreaks: null,
    colBreaks: null
  }, options.pageSetup);

  // using shared strings creates a smaller xlsx file but may use more memory
  this.useSharedStrings = options.useSharedStrings || false;

  this._workbook = options.workbook;

  // views
  this._views = options.views || [];

  // start writing to stream now
  this._writeOpenWorksheet();

  this.startedData = false;
};

WorksheetWriter.prototype = {
  get workbook() {
    return this._workbook;
  },

  get stream() {
    if (!this._stream) {
      this._stream = this._workbook._openStream('/xl/worksheets/sheet' + this.id + '.xml');

      // pause stream to prevent 'data' events
      this._stream.pause();
    }
    return this._stream;
  },

  // destroy - not a valid operation for a streaming writer
  // even though some streamers might be able to, it's a bad idea.
  destroy: function() {
    throw new Error('Invalid Operation: destroy');
  },

  commit: function() {
    if (this.committed) {
      return;
    }
    var self = this;
    // commit all rows
    this._rows.forEach(function(cRow) {
      if (cRow) {
        // write the row to the stream
        self._writeRow(cRow);
      }
    });

    // we _cannot_ accept new rows from now on
    this._rows = null;

    if (!this.startedData) {
        this._writeOpenSheetData();
    }
    this._writeCloseSheetData();
    this._writeMergeCells();

    // for some reason, Excel can't handle dimensions at the bottom of the file
    //this._writeDimensions();

    this._writeHyperlinks();
    this._writeDataValidations();
    this._writePageMargins();
    this._writePageSetup();
    this._writeCloseWorksheet();

    // signal end of stream to workbook
    this.stream.end();

    // also commit the hyperlinks if any
    this._hyperlinkWriter.commit();

    this.committed = true;
  },

  // return the current dimensions of the writer
  get dimensions() {
    return this._dimensions;
  },

  get views() {
    return this._views;
  },

  // =========================================================================
  // Columns

  // get the current columns array.
  get columns() {
    return this._columns;
  },

  // set the columns from an array of column definitions.
  // Note: any headers defined will overwrite existing values.
  set columns(value) {
    var self = this;

    // calculate max header row count
    this._headerRowCount = value.reduce(function(pv,cv) {
      var headerCount = cv.header ? 1 : (cv.headers ? cv.headers.length : 0);
      return Math.max(pv, headerCount)
    }, 0);

    // construct Column objects
    var count = 1;
    var _columns = this._columns = [];
    value.forEach(function(defn) {
      var column = new Column(self, count++, false);
      _columns.push(column);
      column.defn = defn;
    });
  },

  // get a single column by col number. If it doesn't exist, it and any gaps before it
  // are created.
  getColumn: function(c) {
    if (typeof c === 'string'){
      // if it matches a key'd column, return that
      var col = this._keys[c];
      if (col) return col;

      // otherise, assume letter
      c = colCache.l2n(c);
    }
    if (!this._columns) { this._columns = []; }
    if (c > this._columns.length) {
      var n = this._columns.length + 1;
      while (n <= c) {
        this._columns.push(new Column(this, n++));
      }
    }
    return this._columns[c-1];
  },

  // =========================================================================
  // Rows
  get _nextRow() {
    return this._rowZero + this._rows.length;
  },

  // iterate over every uncommitted row in the worksheet, including maybe empty rows
  eachRow: function(options, iteratee) {
    if (!iteratee) {
      iteratee = options;
      options = undefined;
    }
    if (options && options.includeEmpty) {
      var n = this._nextRow;
      for (var i = this._rowZero; i < n; i++) {
        iteratee(this.getRow(i), i);
      }
    } else {
      this._rows.forEach(function(row) {
        if (row.hasValues) {
          iteratee(row, row.number);
        }
      });
    }
  },

  _commitRow: function(cRow) {
    // since rows must be written in order, we commit all rows up till and including cRow
    var found = false;
    while (this._rows.length && !found) {
      var row = this._rows.shift();
      this._rowZero++;
      if (row) {
        this._writeRow(row);
        found = (row.number === cRow.number);
        this._rowZero = row.number + 1;
      }
    }
  },

  get lastRow() {
    // returns last uncommitted row
    if (this._rows.length) {
      return this._rows[this._rows.length-1];
    } else {
      return undefined;
    }
  },

  // find a row (if exists) by row number
  findRow: function(rowNumber) {
    var index = rowNumber - this._rowZero;
    return this._rows[index];
  },

  getRow: function(rowNumber) {
    var index = rowNumber - this._rowZero;

    // may fail if rows have been comitted
    if (index < 0) {
      throw new Error('Out of bounds: this row has been committed');
    }
    var row = this._rows[index];
    if (!row) {
      this._rows[index] = row = new Row(this, rowNumber);
    }
    return row;
  },

  addRow: function(value) {
    var row = new Row(this, this._nextRow);
    this._rows[row.number-this._rowZero] = row;
    row.values = value;
    return row;
  },

  // ================================================================================
  // Cells

  // returns the cell at [r,c] or address given by r. If not found, return undefined
  findCell: function(r, c) {
    var address = colCache.getAddress(r, c);
    var row = this.findRow(address.row);
    return row ? row.findCell(address.column) : undefined;
  },

  // return the cell at [r,c] or address given by r. If not found, create a new one.
  getCell: function(r, c) {
    var address = colCache.getAddress(r, c);
    var row = this.getRow(address.row);
    return row._getCell(address);
  },

  mergeCells: function() {
    // may fail if rows have been comitted
    var dimensions = new Dimensions(Array.prototype.slice.call(arguments, 0)); // convert arguments into Array

    // check cells aren't already merged
    this._merges.forEach(function(merge) {
      if (merge.intersects(dimensions)) {
        throw new Error('Cannot merge alreay merged cells');
      }
    });

    // apply merge
    var master = this.getCell(dimensions.top, dimensions.left);
    for (var i = dimensions.top; i <= dimensions.bottom; i++) {
      for (var j = dimensions.left; j <= dimensions.right; j++) {
        if ((i > dimensions.top) || (j > dimensions.left)) {
          this.getCell(i,j).merge(master);
        }
      }
    }

    // index merge
    this._merges.push(dimensions);
  },

  // ================================================================================
  _write: function(text) {
    // var xml = stringBufStack.pop();
    xml.reset();
    xml.addText(text);
    this.stream.write(xml);
    //stringBufStack.push(xml);
  },
  _writeSheetProperties: function(xml, properties, pageSetup) {
    var sheetPropertiesModel = {
      tabColor: properties && properties.tabColor,
      pageSetup: pageSetup && pageSetup.fitToPage ? {
        fitToPage: pageSetup.fitToPage
      } : undefined
    };

    xml.addText(xform.sheetProperties.toXml(sheetPropertiesModel));
  },
  _writeSheetFormatProperties: function(xml, properties) {
    var sheetFormatPropertiesModel = properties ? {
      defaultRowHeight: properties.defaultRowHeight,
      dyDescent: properties.dyDescent,
      outlineLevelCol: properties.outlineLevelCol,
      outlineLevelRow: properties.outlineLevelRow
    } : undefined;

    xml.addText(xform.sheetFormatProperties.toXml(sheetFormatPropertiesModel));
  },
  _writeOpenWorksheet: function() {
    // var xml = stringBufStack.pop();
    xml.reset();

    xml.addText('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    xml.addText('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">');

    this._writeSheetProperties(xml, this.properties, this.pageSetup);

    xml.addText(xform.sheetViews.toXml(this.views));

    this._writeSheetFormatProperties(xml, this.properties);

    this.stream.write(xml);
    //stringBufStack.push(xml);
  },
  _writeColumns: function() {
    var cols = Column.toModel(this.columns);
    if (cols) {
      xform.columns.prepare(cols, {styles: this._workbook.styles});
      this.stream.write(xform.columns.toXml(cols));
    }
  },
  _writeOpenSheetData: function() {
    this._write('<sheetData>');
  },
  _writeRow: function(row) {
    var self = this;

    if (!this.startedData) {
      this._writeColumns();
      this._writeOpenSheetData();
      this.startedData = true;
    }

    if (row.hasValues || row.height) {
      var model = row.model;
      var options = {
        styles: this._workbook.styles,
        sharedStrings: self.useSharedStrings ? self._workbook.sharedStrings : undefined,
        hyperlinks: this._hyperlinkWriter,
        merges: this._merges
      };
      xform.row.prepare(model, options);
      this.stream.write(xform.row.toXml(model));
    }
  },
  _writeCloseSheetData: function() {
    this._write('</sheetData>');
  },
  _writeMergeCells: function() {

    if (this._merges.length) {
      // var xml = stringBufStack.pop();
      xml.reset();
      xml.addText('<mergeCells count="' + this._merges.length + '">');
      this._merges.forEach(function(merge) {
        xml.addText('<mergeCell ref="' + merge + '"/>');
      });
      xml.addText('</mergeCells>');

      this.stream.write(xml);
      //stringBufStack.push(xml);
    }
  },
  _writeHyperlinks: function() {
    this.stream.write(xform.hyperlinks.toXml(this._hyperlinkWriter._hyperlinks));
  },
  _writeDataValidations: function() {
    this.stream.write(xform.dataValidations.toXml(this.dataValidations.model));
  },
  _writePageMargins: function() {
    this.stream.write(xform.pageMargins.toXml(this.pageSetup.margins));
  },
  _writePageSetup: function() {
    this.stream.write(xform.pageSeteup.toXml(this.pageSetup));
  },
  _writeDimensions: function() {
    this._write('<dimension ref="' + this._dimensions + '"/>');
  },
  _writeCloseWorksheet: function() {
    this._write('</worksheet>');
  }
};