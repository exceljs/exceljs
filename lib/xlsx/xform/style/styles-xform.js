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
var Enums = require('../../../doc/enums');
var XmlStream = require('../../../utils/xml-stream');

var BaseXform = require('../base-xform');
var StaticXform = require('../static-xform');
var ListXform = require('../list-xform');
var FontXform = require('./font-xform');
var FillXform = require('./fill-xform');
var BorderXform = require('./border-xform');
var NumFmtXform = require('./numfmt-xform');
var StyleXform = require('./style-xform');

// custom numfmt ids start here
var NUMFMT_BASE = 164;

// =============================================================================
// StylesXform is used to generate and parse the styles.xml file
// it manages the collections of fonts, number formats, alignments, etc
var StylesXform = module.exports = function(initialise) {
  this.map = {
    numFmts: new ListXform({tag: 'numFmts', count: true, childXform: new NumFmtXform()}),
    fonts: new ListXform({tag: 'fonts', count: true, childXform: new FontXform(), $: {"x14ac:knownFonts": 1}}),
    fills: new ListXform({tag: 'fills', count: true, childXform: new FillXform()}),
    borders: new ListXform({tag: 'borders', count: true, childXform: new BorderXform()}),
    cellStyleXfs: new ListXform({tag: 'cellStyleXfs', count: true, childXform: new StyleXform()}),
    cellXfs: new ListXform({tag: 'cellXfs', count: true, childXform: new StyleXform({xfId: true})}),
    
    // for style manager
    numFmt: new NumFmtXform(),
    font: new FontXform(),
    fill: new FillXform(),
    border:  new BorderXform(),
    style: new StyleXform({xfId: true}),

    cellStyles: StylesXform.STATIC_XFORMS.cellStyles,
    dxfs: StylesXform.STATIC_XFORMS.dxfs,
    tableStyles: StylesXform.STATIC_XFORMS.tableStyles,
    extLst: StylesXform.STATIC_XFORMS.extLst
  };

  if (initialise) {
    // StylesXform also acts as style manager and is used to build up styles-model during worksheet processing
    this.init();
  }
};

utils.inherits(StylesXform, BaseXform, {
  STYLESHEET_ATTRIBUTES: {
    'xmlns': "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    'xmlns:mc': "http://schemas.openxmlformats.org/markup-compatibility/2006",
    'mc:Ignorable': "x14ac x16r2",
    'xmlns:x14ac': "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    'xmlns:x16r2': "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"
  },
  STATIC_XFORMS: {
    cellStyles: new StaticXform({tag: 'cellStyles', $: {count:1}, c: [{tag:'cellStyle', $: {name:'Normal', xfId:0, builtinId:0}}]}),
    dxfs: new StaticXform({tag: 'dxfs', $: {count:0}}),
    tableStyles: new StaticXform({tag: 'tableStyles', $: {count: 0, defaultTableStyle: 'TableStyleMedium2', defaultPivotStyle: 'PivotStyleLight16'}}),
    extLst: new StaticXform({tag: 'extLst', c: [
      {tag: 'ext', $: {uri:'{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}', 'xmlns:x14': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'}, c: [
        {tag: 'x14:slicerStyles', $:{defaultSlicerStyle: 'SlicerStyleLight1'}}
      ]},
      {tag: 'ext', $: {uri:'{9260A510-F301-46a8-8635-F512D64BE5F5}', 'xmlns:x15': 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'}, c: [
        {tag: 'x15:timelineStyles', $:{defaultTimelineStyle: 'TimeSlicerStyleLight1'}}
      ]}
    ]})
  }
},{

  initIndex: function() {
    this.index = {
      style: {},
      numFmt: {},
      numFmtNextId: 164, // start custom format ids here
      font: {},
      border: {},
      fill: {}
    };
  },

  init: function() {
    // Prepare for Style Manager role
    this.model = {
      styles: [],
      numFmts: [],
      fonts: [],
      borders: [],
      fills: []
    };

    this.initIndex();

    // default (zero) font
    this._addFont({size: 11, color: {theme:1}, name: 'Calibri', family:2, scheme:'minor'});

    // default (zero) border
    this._addBorder({});

    // add default (all zero) style
    this._addStyle({numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0});

    // add default fills
    this._addFill({type: 'pattern', pattern:'none'});
    this._addFill({type: 'pattern', pattern:'gray125'});
  },

  render: function(xmlStream, model) {
    model = model || this.model;
    //
    //   <fonts count="2" x14ac:knownFonts="1">
    xmlStream.openXml(XmlStream.StdDocAttributes);

    xmlStream.openNode('styleSheet', StylesXform.STYLESHEET_ATTRIBUTES);

    if (this.index) {
      // model has been built by style manager role (contains xml)
      if (model.numFmts && model.numFmts.length) {
        xmlStream.openNode('numFmts', {count: model.numFmts.length});
        model.numFmts.forEach(function(numFmtXml) {
          xmlStream.writeXml(numFmtXml);
        });
        xmlStream.closeNode();
      }

      xmlStream.openNode('fonts', {count: model.fonts.length});
      model.fonts.forEach(function(fontXml) {
        xmlStream.writeXml(fontXml);
      });
      xmlStream.closeNode();

      xmlStream.openNode('fills', {count: model.fills.length});
      model.fills.forEach(function(fillXml) {
        xmlStream.writeXml(fillXml);
      });
      xmlStream.closeNode();

      xmlStream.openNode('borders', {count: model.borders.length});
      model.borders.forEach(function(borderXml) {
        xmlStream.writeXml(borderXml);
      });
      xmlStream.closeNode();

      this.map.cellStyleXfs.render(xmlStream, [{numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0}]);

      xmlStream.openNode('cellXfs', {count: model.styles.length});
      model.styles.forEach(function(styleXml) {
        xmlStream.writeXml(styleXml);
      });
      xmlStream.closeNode();
    } else {
      // model is plain JSON and needs to be xformed
      this.map.numFmts.render(xmlStream, model.numFmts);
      this.map.fonts.render(xmlStream, model.fonts);
      this.map.fills.render(xmlStream, model.fills);
      this.map.borders.render(xmlStream, model.borders);
      this.map.cellStyleXfs.render(xmlStream, [{numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0}]);
      this.map.cellXfs.render(xmlStream, model.styles);
    }

    // trailing static stuff
    StylesXform.STATIC_XFORMS.cellStyles.render(xmlStream);
    StylesXform.STATIC_XFORMS.dxfs.render(xmlStream);
    StylesXform.STATIC_XFORMS.tableStyles.render(xmlStream);
    StylesXform.STATIC_XFORMS.extLst.render(xmlStream);

    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      switch(node.name) {
        case 'styleSheet':
          this.initIndex();
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
        case 'styleSheet':
          var model = this.model = {};
          var add = function(name, xform) {
            if (xform.model && xform.model.length) {
              model[name] = xform.model;
            }
          };
          add('numFmts', this.map.numFmts);
          add('fonts', this.map.fonts);
          add('fills', this.map.fills);
          add('borders', this.map.borders);
          add('styles', this.map.cellXfs);

          // index numFmts
          this.index = {
            model:[],
            numFmt:  []
          };
          if (model.numFmts) {
            var numFmtIndex = this.index.numFmt;
            model.numFmts.forEach(function (numFmt) {
              numFmtIndex[numFmt.id] = numFmt.formatCode;
            });
          }

          return false;
        default:
          // not quite sure how we get here!
          return true;
      }
    }
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

    var style = {};
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
      style.fontId = this._addFont(model.font);
    }

    if (model.border) {
      style.borderId = this._addBorder(model.border);
    }

    if (model.fill) {
      style.fillId = this._addFill(model.fill);
    }

    if (model.alignment) {
      style.alignment =  model.alignment;
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
    var model = this.index.model[id];
    if (model) return model;

    // if the style doesn't exist return null
    var style = this.model.styles[id];
    if (!style) return null;

    // build a new model
    model = this.index.model[id] = {};

    // -------------------------------------------------------
    // number format
    if (style.numFmtId) {
      var numFmt = this.index.numFmt[style.numFmtId] || NumFmtXform.getDefaultFmtCode(style.numFmtId);
      if (numFmt) {
        model.numFmt = numFmt;
      }
    }

    function addStyle(name, group, id) {
      if (id) {
        var part = group[id];
        if (part) {
          model[name] = part;
        }
      }
    }

    addStyle('font', this.model.fonts, style.fontId);
    addStyle('border', this.model.borders, style.borderId);
    addStyle('fill', this.model.fills, style.fillId);

    // -------------------------------------------------------
    // alignment
    if (style.alignment) {
      model.alignment = style.alignment;
    }

    return model;
  },

  // =========================================================================
  // Private Interface
  _addStyle: function(style) {
    var xml = this.map.style.toXml(style);
    var index = this.index.style[xml];
    if (index === undefined) {
      index = this.index.style[xml] = this.model.styles.length;
      this.model.styles.push(xml);
    }
    return index;
  },

  // =========================================================================
  // Number Formats
  _addNumFmtStr: function(formatCode) {
    // check if default format
    var index = NumFmtXform.getDefaultFmtId(formatCode);
    if (index !== undefined) return index;

    // check if already in
    index = this.index.numFmt[formatCode];
    if (index !== undefined) return index;


    index = this.index.numFmt[formatCode] = NUMFMT_BASE + this.model.numFmts.length;
    var xml = this.map.numFmt.toXml({id: index, formatCode: formatCode});
    this.model.numFmts.push(xml);
    return index;
  },

  // =========================================================================
  // Fonts
  _addFont: function(font) {
    var xml = this.map.font.toXml(font);
    var index = this.index.font[xml];
    if (index === undefined) {
      index = this.index.font[xml] = this.model.fonts.length;
      this.model.fonts.push(xml);
    }
    return index;
  },

  // =========================================================================
  // Borders
  _addBorder: function(border) {
    var xml = this.map.border.toXml(border);
    var index = this.index.border[xml];
    if (index === undefined) {
      index = this.index.border[xml] = this.model.borders.length;
      this.model.borders.push(xml);
    }
    return index;
  },

  // =========================================================================
  // Fills
  _addFill: function(fill) {
    var xml = this.map.fill.toXml(fill);
    var index = this.index.fill[xml];
    if (index === undefined) {
      index = this.index.fill[xml] = this.model.fills.length;
      this.model.fills.push(xml);
    }
    return index;
  }

  // =========================================================================
});

// the stylemanager mock acts like StyleManager except that it always returns 0 or {}
StylesXform.Mock = function() {
  StylesXform.call(this);
  this.model = {
    styles: [{numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0}],
    numFmts: [],
    fonts: [{size: 11, color: {theme:1}, name: 'Calibri', family:2, scheme:'minor'}],
    borders: [{}],
    fills:  [
      {type: 'pattern', pattern:'none'},
      {type: 'pattern', pattern:'gray125'}
    ]
  };
};

utils.inherits(StylesXform.Mock, StylesXform, {
  // =========================================================================
  // Style Manager Interface

  // override normal behaviour - consume and dispose
  parseStream: function(stream) {
    stream.autodrain();
    return Promise.resolve();
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
      var dateStyle = {
        numFmtId: NumFmtXform.getDefaultFmtId('mm-dd-yy')
      };
      this._dateStyleId = this.model.styles.length;
      this.model.styles.push(dateStyle);
    }
    return this._dateStyleId;
  },

  // given a styleId (i.e. s="n"), get the cell's style model
  // objects are shared where possible.
  getStyleModel: function(/*id*/) {
    return {}
  }
});
