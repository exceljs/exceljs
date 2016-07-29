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
"use strict";

var Enums = require("../../../doc/enums");

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var valid = {
  horizontalValues: [
    "left",
    "center",
    "right",
    "fill",
    "centerContinuous",
    "distributed",
    "justify"
  ].reduce(function(p,v){p[v]=true; return p;}, {}),
  horizontal: function(value) {
    return this.horizontalValues[value] ? value : undefined;
  },

  verticalValues: [
    "top",
    "middle",
    "bottom",
    "distributed",
    "justify"
  ].reduce(function(p,v){p[v]=true; return p;}, {}),
  vertical: function(value) {
    if (value === "middle") return "center";
    return this.verticalValues[value] ? value : undefined;
  },
  wrapText: function(value) {
    return value ? true : undefined;
  },
  shrinkToFit: function(value) {
    return value ? true : undefined;
  },
  textRotation: function(value) {
    switch(value) {
      case "vertical":
        return value;
      default:
        value = utils.validInt(value);
        return (value >= -90) && (value <= 90) ? value : undefined;
    }
  },
  indent: function(value) {
    value = utils.validInt(value);
    return Math.max(0, value);
  },
  readingOrderValues: [
    {k:"r2l",v:1},
    {k:"l2r",v:2},
    {k:Enums.ReadingOrder.RightToLeft,v:1},
    {k:Enums.ReadingOrder.LeftToRight,v:2}
  ].reduce(function(p,v) {p[v.k]=v.v; return p;}, {}),

  readingOrder: function(value) {
    return this.readingOrderValues[value];
  }
};

var textRotationXform = {
  toXml:  function(textRotation) {
    textRotation = valid.textRotation(textRotation);
    if (textRotation) {
      if (textRotation == "vertical") {
        return 255;
      } else {
        var tr = Math.round(textRotation);
        if ((tr >= 0) && (tr <= 90)) {
          return tr;
        } else if ((tr < 0) && (tr >= -90)) {
          return 90 - tr;
        }
      }
    }
  },
  toModel: function(textRotation) {
    var tr = utils.validInt(textRotation);
    if (tr !== undefined) {
      if (tr == 255) {
        return "vertical";
      } else if((tr >= 0) && (tr <= 90)) {
        return tr;
      } else if ((tr > 90) && (tr <= 180)) {
        return 90 - tr;
      }
    }
  }
};

// Alignment encapsulates translation from style.alignment model to/from xlsx
var AlignmentXform = module.exports = function() {
};

utils.inherits(AlignmentXform, BaseXform, {

  get tag() { return 'alignment'; },

  render: function(xmlStream, model) {

    xmlStream.addRollback();
    xmlStream.openNode('alignment');

    var isValid = false;
    function add(name, value) {
      if (value) {
        xmlStream.addAttribute(name, value);
        isValid = true;
      }
    }
    add('horizontal', valid.horizontal(model.horizontal));
    add('vertical', valid.vertical(model.vertical));
    add('wrapText', valid.wrapText(model.wrapText) ? '1' : false);
    add('shrinkToFit', valid.shrinkToFit(model.shrinkToFit) ? '1' : false);
    add('indent', valid.indent(model.indent));
    add('textRotation', textRotationXform.toXml(model.textRotation));
    add('readingOrder', valid.readingOrder(model.readingOrder));

    xmlStream.closeNode();

    if (isValid) {
      xmlStream.commit();
    } else {
      xmlStream.rollback();
    }
  },
  
  parseOpen: function(node) {
    var model = {};
    
    var valid = false;
    function add(truthy, name, value) {
      if (truthy) {
        model[name] = value;
        valid = true;
      }
    }
    add(node.attributes.horizontal, 'horizontal', node.attributes.horizontal);
    add(node.attributes.vertical, 'vertical', node.attributes.vertical === 'center' ? 'middle' : node.attributes.vertical);
    add(node.attributes.wrapText, 'wrapText', node.attributes.wrapText ? true : false);
    add(node.attributes.shrinkToFit, 'shrinkToFit', node.attributes.shrinkToFit ? true : false);
    add(node.attributes.indent, 'indent', parseInt(node.attributes.indent));
    add(node.attributes.textRotation, 'textRotation', textRotationXform.toModel(node.attributes.textRotation));
    add(node.attributes.readingOrder, 'readingOrder', node.attributes.readingOrder);
    
    this.model = valid ? model : null;
  },
  
  parseText: function() {
  },
  parseClose: function() {
    return false;
  }
});
