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

var ColorXform = require('./color-xform');
var BooleanXform = require('../simple/boolean-xform');
var IntegerXform = require('../simple/integer-xform');
var StringXform = require('../simple/string-xform');
var UnderlineXform = require('./underline-xform');

var _ = require('../../../utils/under-dash');
var utils  = require('../../../utils/utils');
var BaseXform = require('../base-xform');

// Font encapsulates translation from font model to xlsx
var FontXform = module.exports = function(options) {
  this.options = options || FontXform.OPTIONS;

  this.map = {
    b: { prop: 'bold', xform: new BooleanXform({tag: 'b', attr: 'val'}) },
    i: { prop: 'italic', xform: new BooleanXform({tag: 'i', attr: 'val'}) },
    u: { prop: 'underline', xform: new UnderlineXform() },
    charset: { prop: 'charset', xform: new IntegerXform({tag: 'charset', attr: 'val'}) },
    color: { prop: 'color', xform: new ColorXform() },
    condense: { prop: 'condense', xform: new BooleanXform({tag: 'condense', attr: 'val'}) },
    extend: { prop: 'extend', xform: new BooleanXform({tag: 'extend', attr: 'val'}) },
    family: { prop: 'family', xform: new IntegerXform({tag: 'family', attr: 'val'}) },
    outline: { prop: 'outline', xform: new BooleanXform({tag: 'outline', attr: 'val'}) },
    scheme: { prop: 'scheme', xform: new StringXform({tag: 'scheme', attr: 'val'}) },
    shadow: { prop: 'shadow', xform: new BooleanXform({tag: 'shadow', attr: 'val'}) },
    strike: { prop: 'strike', xform: new BooleanXform({tag: 'strike', attr: 'val'}) },
    sz: { prop: 'size', xform: new IntegerXform({tag: 'sz', attr: 'val'}) }
  };
  this.map[this.options.fontNameTag] = { prop: 'name', xform: new StringXform({tag: this.options.fontNameTag, attr: 'val'}) };
};

FontXform.OPTIONS = {
  tagName: 'font',
  fontNameTag: 'name'
};

utils.inherits(FontXform, BaseXform, {

  get tag() { return this.options.tagName; },

  render: function(xmlStream, model) {
    var map =  this.map;

    xmlStream.openNode(this.options.tagName);
    _.each(this.map, function(defn, tag) {
      map[tag].xform.render(xmlStream, model[defn.prop]);
    });
    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else if(this.map[node.name]) {
      this.parser = this.map[node.name].xform;
      return this.parser.parseOpen(node);
    } else {
      switch (node.name) {
        case this.options.tagName:
          this.model = {};
          return true;
        default:
          return false;
      }
    }
  },
  parseText:  function(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },
  parseClose: function(name) {
    if (this.parser && !this.parser.parseClose(name)) {
      var item = this.map[name];
      if (this.parser.model) {
        this.model[item.prop] = this.parser.model;
      }
      this.parser  = undefined;
      return true;
    } else {
      switch (name) {
        case this.options.tagName:
          return false;
        default:
          return true;
      }
    }
  }
});
