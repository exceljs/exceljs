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

var TextXform = require('./text-xform');
var RichTextXform = require('./rich-text-xform');
var PhoneticTextXform = require('./phonetic-text-xform');

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

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
    t: new TextXform(),
    rPh: new PhoneticTextXform()
  };
};


utils.inherits(SharedStringXform, BaseXform, {

  get tag() { return 'si'; },

  render: function(xmlStream, model) {
    xmlStream.openNode(this.tag);
    if (model && model.hasOwnProperty('richText') && model.richText) {
      var r = this.map.r;
      model.richText.forEach(function(text) {
        r.render(xmlStream, text);
      });
    } else if (model !== undefined && model !== null) {
      this.map.t.render(xmlStream, model);
    }
    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    var name = node.name;
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else if (name === this.tag) {
      this.model = {};
      return true;
    } else {
      this.parser = this.map[name];
      if (this.parser) {
        this.parser.parseOpen(node);
        return true;
      }
      return false;
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
            var rt = this.model.richText;
            if (!rt) { rt = this.model.richText = []; }
            rt.push(this.parser.model);
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
        case this.tag:
          return false;
        default:
          return true;
      }
    }
  }
});
