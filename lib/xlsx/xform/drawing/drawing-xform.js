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

var utils = require('../../../utils/utils');
var XmlStream = require('../../../utils/xml-stream');

var BaseXform = require('../base-xform');
var TwoCellAnchorXform = require('./two-cell-anchor-xform');

var WorkSheetXform = module.exports = function() {
  this.map = {
    'xdr:twoCellAnchor':  new TwoCellAnchorXform()
  };
};

utils.inherits(WorkSheetXform, BaseXform, {
  DRAWING_ATTRIBUTES: {
    'xmlns:xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
  }
},{
  get tag() { return 'xdr:wsDr'; },

  prepare: function(model, options) {
    var twoCellAnchorXform = this.map['xdr:twoCellAnchor'];
    model.anchors.forEach(function(item, index) {
      twoCellAnchorXform.prepare(item, {index: index});
    });
  },

  render: function(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode(this.tag, WorkSheetXform.DRAWING_ATTRIBUTES);

    var twoCellAnchorXform = this.map['xdr:twoCellAnchor'];
    model.anchors.forEach(function(item) {
      twoCellAnchorXform.render(xmlStream, item);
    });

    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          anchors: [],
        };
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  },

  parseText: function(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },

  parseClose: function(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model.anchors.push(this.parser.model);
        this.parser = undefined;
      }
      return true;
    }
    switch(name) {
      case this.tag:
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  },

  reconcile: function(model, options) {
    model.anchors.forEach(anchor => {
      this.map['xdr:twoCellAnchor'].reconcile(anchor, options);
    });
  }
});
