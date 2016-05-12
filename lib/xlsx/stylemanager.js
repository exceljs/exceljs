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
var _ = require('underscore');
var Bluebird = require('bluebird');
var Sax = require('sax');
var Style = require('./style');
var NumFmt = require('./numfmt');
var Alignment = require('./alignment');
var Border = require('./border');
var Fill = require('./fill');

var FontXform = require('./xform/font-xform');

var utils = require('../utils/utils');
var Enums = require('../doc/enums');

// =============================================================================
// StyleManager is used to generate and parse the styles.xml file
// it manages the collections of fonts, number formats, alignments, etc
var StyleManager = module.exports = function() {
  this.styles = [];
  this.stylesIndex = {};

  this.models = []; // styleId -> model

  this.numFmts = [];
  this.numFmtHash = {}; //numFmtId -> numFmt
  this.numFmtNextId = 164; // start custom format ids here

  this.fonts = []; // array of font
  this.fontIndex = {}; // hash of xml->fontId

  this.borders = []; // array of border
  this.borderIndex = {}; // hash of xml->borderId

  this.fills = []; // array of Fill
  this.fillIndex = {}; // hash of xml->fillId

  // Note: aligments stored inside the style objects

  // ---------------------------------------------------------------
  // Defaults

  // default (zero) font
  this._addFont(new FontXform({ size: 11, color: {theme:1}, name: 'Calibri', family:2, scheme:'minor'}));

  // default (zero) border
  this._addBorder(new Border());

  // add default (all zero) style
  this._addStyle(new Style());

  // add default fills
  this._addFill(new Fill({pattern:'none'}));
  this._addFill(new Fill({pattern:'gray125'}));

};

utils.inherits(StyleManager, events.EventEmitter, {
  // =========================================================================
  // Public Interface

  // addToZip - using styles.xml template, add the styles.xml file to the zip
  addToZip: function(zip) {
    var self = this;
    return utils.fetchTemplate(require.resolve('./styles.xml'))
      .then(function(template){
        return template(self);
      })
      .then(function(data) {
        zip.append(data, { name: '/xl/styles.xml' });
        return zip;
      });
  },

  // parse the styles.xml file pulling out fonts, numFmts, etc
  parse: function(stream) {
    var self = this;
    var parser = Sax.createStream(true, {});

    var inNumFmts = false;
    var inCellXfs = false;
    var inFonts = false;
    var inBorders = false;
    var inFills = false;

    var style = null;
    var font = null;
    var border = null;
    var fill = null;

    return new Bluebird(function(resolve, reject) {

      parser.on('opentag', function(node) {
        if (inNumFmts) {
          if (node.name == 'numFmt') {
            var numFmt = new NumFmt();
            numFmt.parse(node);
            self._addNumFmt(numFmt);
          }
        } else if (inCellXfs) {
          if (style) {
            style.parse(node);
          } else if (node.name == 'xf') {
            style = new Style();
            style.parse(node);
          }
        } else if (inFonts) {
          if (font) {
            font.parseOpen(node);
          } else if (node.name == 'font') {
            font = new FontXform();
            font.parseOpen(node);
          }
        } else if (inBorders) {
          if (node.name == 'border') {
            border = new Border();
          }
          border.parse(node);
        } else if (inFills) {
          if (node.name == 'fill') {
            fill = new Fill();
          }
          fill.parse(node);
        } else {
          switch(node.name) {
            case 'cellXfs':
              inCellXfs = true;
              break;
            case 'numFmts':
              inNumFmts = true;
              break;
            case 'fonts':
              inFonts = true;
              break;
            case 'borders':
              inBorders = true;
              break;
            case 'fills':
              inFills = true;
              break;
          }
        }
      });
      parser.on('closetag', function (name) {
        if (inCellXfs) {
          switch(name) {
            case 'cellXfs':
              inCellXfs = false;
              break;
            case 'xf':
              self._addStyle(style);
              style = null;
              break;
          }
        } else if (inNumFmts) {
          switch(name) {
            case 'numFmts':
              inNumFmts = false;
              break;
          }
        } else if (inFonts) {
          switch(name) {
            case 'fonts':
              inFonts = false;
              break;
            case 'font':
              self._addFont(font);
              font = null;
              break;
            default:
              font.parseClose(name);
              break;
          }
        } else if (inBorders) {
          switch(name) {
            case 'borders':
              inBorders = false;
              break;
            case 'border':
              self._addBorder(border);
              border = null;
              break;
          }
        } else if (inFills) {
          switch(name) {
            case 'fills':
              inFills = false;
              break;
            case 'fill':
              self._addFill(fill);
              fill = null;
              break;
          }
        }
      });
      parser.on('end', function() {
        // warning: if style, font, border, fill, inBorders, inFills, inFonts, inCellXfs, inNumFmts are true!
        resolve(self);
      });
      parser.on('error', function (error) {
        reject(null,error);
      });
      stream.pipe(parser);
    });
  },

  // add a cell's style model to the collection
  // each style property is processed and cross-referenced, etc.
  // the styleId is returned. Note: cellType is used when numFmt not defined
  addStyleModel: function(model, cellType) {
    if (!model) {
      return 0;
    }

    // if we have seen this style object before, assume it has the same styleId
    if (this.weakMap && this.weakMap.has(model)) {
      return this.weakMap.get(model);
    }

    var style = new Style();
    cellType = cellType || Enums.ValueType.Number;

    if (model.numFmt) {
      style.numFmtId = this._addNumFmtStr(model.numFmt);
    } else {
      switch(cellType) {
        case Enums.ValueType.Number:
          style.numFmtId = this._addNumFmtStr('General');
          break;
        case Enums.ValueType.Date:
          style.numFmtId = this._addNumFmtStr('mm-dd-yy');
          break;
      }
    }

    if (model.font) {
      style.fontId = this._addFont(new FontXform(model.font));
    }

    if (model.border) {
      style.borderId = this._addBorder(new Border(model.border));
    }

    if (model.fill) {
      style.fillId = this._addFill(new Fill(model.fill));
    }

    if (model.alignment) {
      style.alignment = new Alignment(model.alignment);
    }

    var styleId = this._addStyle(style);
    if (this.weakMap) {
      this.weakMap.set(model, styleId);
    }
    return styleId;
  },

  // given a styleId (i.e. s="n"), get the cell's style model
  // objects are shared where possible.
  getStyleModel: function(id) {
    // have we built this model before?
    var model = this.models[id];
    if (model) return model;

    // if the style doesn't exist return null
    var style = this.styles[id];
    if (!style) return null;

    // build a new model
    var model = this.models[id] = {};

    // -------------------------------------------------------
    // number format
    var numFmt = (this.numFmtHash[style.numFmtId] && this.numFmtHash[style.numFmtId].formatCode) ||
      NumFmt.getDefaultFmtCode(style.numFmtId);
    if (numFmt) {
      model.numFmt = numFmt;
    }

    // -------------------------------------------------------
    // font
    var font = this.fonts[style.fontId];
    if (font) {
      model.font = font.model
    }

    // -------------------------------------------------------
    // border
    var border = this.borders[style.borderId];
    if (border) {
      model.border = border.model
    }

    // -------------------------------------------------------
    // fill
    var fill = this.fills[style.fillId];
    if (fill) {
      model.fill = fill.model
    }

    // -------------------------------------------------------
    // alignment
    if (style.alignment) {
      model.alignment = style.alignment.model;
    }

    return model;
  },

  // =========================================================================
  // Private Interface
  _addStyle: function(style) {
    var index = this.stylesIndex[style.xml];
    if (index === undefined) {
      this.styles.push(style);
      index = this.stylesIndex[style.xml] = this.styles.length - 1;
    }
    return index;
  },

  // =========================================================================
  // Number Formats
  _addNumFmtStr: function(formatCode) {
    // check if default format
    var index = NumFmt.getDefaultFmtId(formatCode);
    if (index !== undefined) return parseInt(index);

    var numFmt = this.numFmtHash[formatCode];
    if (numFmt) return numFmt.id;

    // search for an unused index
    index = this.numFmtNextId++;
    return this._addNumFmt(new NumFmt(index, formatCode));
  },
  _addNumFmt: function(numFmt) {
    // called during parse
    this.numFmts.push(numFmt);
    this.numFmtHash[numFmt.id] = numFmt;
    return numFmt.id;
  },

  // =========================================================================
  // Fonts
  _addFont: function(font) {
    var fontId = this.fontIndex[font.xml];
    if (fontId === undefined) {
      fontId = this.fontIndex[font.xml] = this.fonts.length;
      this.fonts.push(font);
    }
    return fontId;
  },

  // =========================================================================
  // Borders
  _addBorder: function(border) {
    var borderId = this.borderIndex[border.xml];
    if (borderId === undefined) {
      borderId = this.borderIndex[border.xml] = this.borders.length;
      this.borders.push(border);
    }
    return borderId;
  },

  // =========================================================================
  // Fills
  _addFill: function(fill) {
    var fillId = this.fillIndex[fill.xml];
    if (fillId === undefined) {
      fillId = this.fillIndex[fill.xml] = this.fills.length;
      this.fills.push(fill);
    }
    return fillId;
  }

  // =========================================================================
});

// the stylemanager mock acts like StyleManager except that it always returns 0 or {}
StyleManager.Mock = function() {
  this.styles = [
    new Style()
  ];
  this._dateStyleId = 0;

  this.numFmts = [];

  this.fonts = [
    new FontXform({ size: 11, color: {theme:1}, name: 'Calibri', family:2, scheme:'minor'})
  ];

  this.borders = [
    new Border()
  ];

  this.fills = [
    new Fill({pattern:'none'}),
    new Fill({pattern:'gray125'})
  ];
};

utils.inherits(StyleManager.Mock, events.EventEmitter, {
  // =========================================================================
  // Public Interface

  // addToZip - using styles.xml template, add the styles.xml file to the zip
  addToZip: StyleManager.prototype.addToZip,

  // parse the styles.xml file pulling out fonts, numFmts, etc
  parse: function(stream) {
    stream.autodrain();
    return Bluebird.resolve();
  },

  // add a cell's style model to the collection
  // each style property is processed and cross-referenced, etc.
  // the styleId is returned. Note: cellType is used when numFmt not defined
  addStyleModel: function(model, cellType) {
    switch (cellType) {
      case Enums.ValueType.Date:
        return this.dateStyleId;
      default:
        return 0;
    }
  },

  get dateStyleId() {
    if (!this._dateStyleId) {
      var dateStyle = new Style();
      dateStyle.numFmtId = NumFmt.getDefaultFmtId('mm-dd-yy');
      this._dateStyleId = this.styles.length;
      this.styles.push(dateStyle);
    }
    return this._dateStyleId;
  },

  // given a styleId (i.e. s="n"), get the cell's style model
  // objects are shared where possible.
  getStyleModel: function(id) {
    return {}
  }
});
