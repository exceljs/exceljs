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

var _ = require('underscore');

var TextXform = require('./text-xform');
var RichTextXform = require('./rich-text-xform');

var utils = require('../../utils/utils');
var BaseXform = require('./base-xform');

// <si>
//   <r></r><r></r>...
// </si>
// <si>
//   <t></t>
// </si>

var SharedStringXform = module.exports = function(model) {
  this.model = model;
  
  this.map = {
    r: new RichTextXform(),
    t: new TextXform()
  };
};


utils.inherits(SharedStringXform, BaseXform, {
  
  write: function(xmlStream, model) {
    model = model || this.model;
    
    if (typeof model === 'undefined') model = '';
    xmlStream.openNode('si');
    if (model.richText) {
      var r = this.map.r;
      _.each(model.richText, function(text) {
        r.write(xmlStream, text);
      });
    } else if (model) {
      this.map.t.write(xmlStream, model);
    }
    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    var name = node.name;
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      switch(name) {
        case 'si':
          this.model = {};
          return true;
        case 'r':
          this.model.richText = this.model.richText || [];
          this.parser  = this.map.r;
          this.parser.parseOpen(node);
          return true;
        case 't':
          this.parser = this.map.t;
          this.parser.parseOpen(node);
          return true;
        default:
          return false;
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
        switch(name) {
          case 'r':
            this.model.richText.push(this.parser.model);
            break;
          case 't':
            this.model = this.parser.model;
            break;
        }
        this.parser = undefined;
      }
      return true;
    } else {
      switch(name) {
        case 'si':
          return false;
        default:
          return true;
      }
    }
  }
});
