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

var _ = require('underscore');
var utils = require('../../utils/utils');
var BaseXform = require('./base-xform');

var CompositeXform = module.exports = function(options) {
  this.tag = options.tag;
  this.attrs = options.attrs;
  this.children = options.children;
  this.map = this.children.reduce(function(map, child) {
    map[child.tag] = child;
    child.name = child.name || child.tag;
    return map;
  });
};

utils.inherits(CompositeXform, BaseXform, {
  prepare: function(model, options) {
    _.each(this.children, function (child) {
      child.xform.prepare(model[child.tag], options);
    });
  },
  
  render: function(xmlStream, model) {
    xmlStream.openNode(this.tag, this.attrs);
    _.each(this.children, function (child) {
      child.xform.render(xmlStream, model[child.name]);
    });
    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.xform.parseOpen(node);
      return true;
    } else {
      switch(node.name) {
        case this.tag:
          this.model = {};
          return true;
        default:
          this.parser = this.map[node.name];
          if (this.parser) {
            this.parser.xform.parseOpen(node);
            return true;
          }
      }
      return false;
    }
  },
  parseText: function(text) {
    if (this.parser) {
      this.parser.xform.parseText(text);
    }
  },
  parseClose: function(name) {
    if (this.parser) {
      if (!this.parser.xform.parseClose(name)) {
        this.model[this.parser.name] = this.parser.xform.model;
        this.parser = undefined;
      }
      return true;
    } else {
      return false;
    }
  },
  reconcile: function(model, options) {
    _.each(this.children, function (child) {
      child.xform.prepare(model[child.tag], options);
    });
  }
});
