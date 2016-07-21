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
var ColorXform = require('../style/color-xform');

var SheetPropertiesXform = module.exports = function() {
  this.map = {
    tabColor: new ColorXform('tabColor')
  }
};

utils.inherits(SheetPropertiesXform, BaseXform, {
  
  render: function(xmlStream, model) {
    if (model.tabColor) {
      xmlStream.openNode('sheetPr');
      this.map.tabColor.render(xmlStream, model.tabColor);
      xmlStream.closeNode();
    }
  },
  
  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else if (node.name === 'sheetPr') {
      return true;
    } else if (this.map[node.name]) {
      this.parser = this.map[node.name];
      this.parser.parseOpen(node);
      return true;
    } else {
      return false;
    }
  },
  parseText: function(text) {
    if (this.parser) {
      this.parser.parseText(text);
      return true;
    }
  },
  parseClose: function(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    } else {
      this.model = {
        tabColor: this.map.tabColor.model
      };
      return false;
    }
  }
});
