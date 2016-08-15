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

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

var WorkbookViewXform = module.exports = function() {
};

utils.inherits(WorkbookViewXform, BaseXform, {
  // 		<workbookView visibility="hidden" xWindow="0" yWindow="0" windowWidth="23040" windowHeight="8808"/>

  render: function(xmlStream, model) {
    var attributes = {
      xWindow: model.x || 0,
      yWindow: model.y || 0,
      windowWidth: model.width || 12000,
      windowHeight: model.height || 24000,
      firstSheet: model.firstSheet,
      activeTab: model.activeTab
    };
    if (model.visibility && model.visibility !== 'visible') {
      attributes.visibility = model.visibility;
    }
    xmlStream.leafNode('workbookView', attributes);
  },
  
  parseOpen: function(node) {
    if (node.name === 'workbookView') {
      var model = this.model = {};
      var addS = function(name, value, dflt) {
        var s = (value !== undefined) ? model[name] = value : dflt;
        if (s !== undefined) {
          model[name] = s;
        }
      };
      var addN = function(name, value, dflt) {
        var n = (value !== undefined) ? model[name] = parseInt(value) : dflt;
        if (n !== undefined) {
          model[name] = n;
        }
      };
      addN('x', node.attributes.xWindow, 0);
      addN('y', node.attributes.yWindow,  0);
      addN('width', node.attributes.windowWidth, 25000);
      addN('height', node.attributes.windowHeight, 10000);
      addS('visibility', node.attributes.visibility, 'visible');
      addN('activeTab', node.attributes.activeTab, undefined);
      addN('firstSheet', node.attributes.firstSheet, undefined);
      return true;
    } else {
      return false;
    }
  },
  parseText: function() {
  },
  parseClose: function() {
    return false;
  }
});
