/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
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

var SharedStringXform = module.exports = function (model) {
  this.model = model;

  this.map = {
    r: new RichTextXform(),
    t: new TextXform(),
    rPh: new PhoneticTextXform()
  };
};

utils.inherits(SharedStringXform, BaseXform, {

  get tag() {
    return 'si';
  },

  render: function render(xmlStream, model) {
    xmlStream.openNode(this.tag);
    if (model && model.hasOwnProperty('richText') && model.richText) {
      var r = this.map.r;
      model.richText.forEach(function (text) {
        r.render(xmlStream, text);
      });
    } else if (model !== undefined && model !== null) {
      this.map.t.render(xmlStream, model);
    }
    xmlStream.closeNode();
  },

  parseOpen: function parseOpen(node) {
    var name = node.name;
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else if (name === this.tag) {
      this.model = {};
      return true;
    }
    this.parser = this.map[name];
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    return false;
  },
  parseText: function parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  },
  parseClose: function parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        switch (name) {
          case 'r':
            var rt = this.model.richText;
            if (!rt) {
              rt = this.model.richText = [];
            }
            rt.push(this.parser.model);
            break;
          case 't':
            this.model = this.parser.model;
            break;
          default:
            break;
        }
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        return true;
    }
  }
});
//# sourceMappingURL=shared-string-xform.js.map
