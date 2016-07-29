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

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var ColXform = module.exports = function() {
};

utils.inherits(ColXform, BaseXform, {

  get tag() { return 'col'; },

  prepare: function(model, options) {
    var styleId = options.styles.addStyleModel(model.style || {});
    if (styleId) {
      model.styleId = styleId;
    }
  },
  
  render: function(xmlStream, model) {
    xmlStream.openNode('col');
    xmlStream.addAttribute('min',  model.min);
    xmlStream.addAttribute('max',  model.max);
    if (model.width) {
      xmlStream.addAttribute('width', model.width);
    }
    if (model.styleId) {
      xmlStream.addAttribute('style',  model.styleId);
    }
    if (model.hidden) {
      xmlStream.addAttribute('hidden',  '1');
    }
    if (model.bestFit) {
      xmlStream.addAttribute('bestFit',  '1');
    }
    if (model.outlineLevel) {
      xmlStream.addAttribute('outlineLevel',  model.outlineLevel);
    }
    if (model.collapsed) {
      xmlStream.addAttribute('collapsed',  '1');
    }
    xmlStream.addAttribute('customWidth',  '1');
    xmlStream.closeNode();
  },
  
  parseOpen: function(node) {
    if (node.name === 'col') {
      var model = this.model = {
        min: parseInt(node.attributes.min || '0'),
        max: parseInt(node.attributes.max || '0'),
        width: parseFloat(node.attributes.width || '0')
      };
      if (node.attributes.style) {
        model.styleId = parseInt(node.attributes.style);
      }
      if (node.attributes.hidden) {
        model.hidden = true;
      }
      if (node.attributes.bestFit) {
        model.bestFit = true;
      }
      if (node.attributes.outlineLevel) {
        model.outlineLevel = parseInt(node.attributes.outlineLevel);
      }
      if (node.attributes.collapsed) {
        model.collapsed = true;
      }
      return true;
    } else {
      return false;
    }
  },
  parseText: function() {
  },
  parseClose: function() {
    return false;
  },
  
  reconcile: function(model, options) {
    // reconcile column styles
    if (model.styleId) {
      model.style = options.styles.getStyleModel(model.styleId);
    }
  }
});
