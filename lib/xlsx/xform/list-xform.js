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

var utils = require('../../utils/utils');
var BaseXform = require('./base-xform');

var ListXform = module.exports = function(options) {
  this.tag = options.tag;
  this.count = options.count;
  this.empty = options.empty;
  this.$count = options.$count || 'count';
  this.$ = options.$;
  this.childXform = options.childXform;
};

utils.inherits(ListXform, BaseXform, {
  prepare: function(model, options) {
    var childXform = this.childXform;
    if (model) {
      model.forEach(function (childModel) {
        childXform.prepare(childModel, options);
      });
    }
  },
  
  render: function(xmlStream, model) {
    if (model && model.length) {
      xmlStream.openNode(this.tag, this.$);
      if (this.count) {
        xmlStream.addAttribute(this.$count, model.length);
      }

      var childXform = this.childXform;
      model.forEach(function (childModel) {
        childXform.render(xmlStream, childModel);
      });

      xmlStream.closeNode();
    } else if (this.empty) {
      xmlStream.leafNode(this.tag);
    }
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      switch(node.name) {
        case this.tag:
          this.model = [];
          return true;
        default:
          if (this.childXform.parseOpen(node)) {
            this.parser = this.childXform;
            return true;
          } else {
            return false;
          }
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
        this.model.push(this.parser.model);
        this.parser = undefined;
      }
      return true;
    } else {
      return false;
    }
  },
  reconcile: function(model, options) {
    if (model) {
      var childXform = this.childXform;
      model.forEach(function (childModel) {
        childXform.reconcile(childModel, options);
      });
    }
  }
});
