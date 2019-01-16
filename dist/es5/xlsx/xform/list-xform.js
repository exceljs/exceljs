/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../utils/utils');
var BaseXform = require('./base-xform');

var ListXform = module.exports = function (options) {
  this.tag = options.tag;
  this.count = options.count;
  this.empty = options.empty;
  this.$count = options.$count || 'count';
  this.$ = options.$;
  this.childXform = options.childXform;
  this.maxItems = options.maxItems;
};

utils.inherits(ListXform, BaseXform, {
  prepare: function prepare(model, options) {
    var childXform = this.childXform;
    if (model) {
      model.forEach(function (childModel) {
        childXform.prepare(childModel, options);
      });
    }
  },

  render: function render(xmlStream, model) {
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

  parseOpen: function parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.model = [];
        return true;
      default:
        if (this.childXform.parseOpen(node)) {
          this.parser = this.childXform;
          return true;
        }
        return false;
    }
  },
  parseText: function parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },
  parseClose: function parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model.push(this.parser.model);
        this.parser = undefined;

        if (this.maxItems && this.model.length > this.maxItems) {
          throw new Error('Max ' + this.childXform.tag + ' count exceeded');
        }
      }
      return true;
    }
    return false;
  },
  reconcile: function reconcile(model, options) {
    if (model) {
      var childXform = this.childXform;
      model.forEach(function (childModel) {
        childXform.reconcile(childModel, options);
      });
    }
  }
});
//# sourceMappingURL=list-xform.js.map
