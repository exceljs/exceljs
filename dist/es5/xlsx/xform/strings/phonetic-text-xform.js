/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var TextXform = require('./text-xform');
var RichTextXform = require('./rich-text-xform');

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

// <rPh sb="0" eb="1">
//   <t>(its pronounciation in KATAKANA)</t>
// </rPh>

var PhoneticTextXform = module.exports = function () {
  this.map = {
    r: new RichTextXform(),
    t: new TextXform()
  };
};

utils.inherits(PhoneticTextXform, BaseXform, {

  get tag() {
    return 'rPh';
  },

  render: function render(xmlStream, model) {
    xmlStream.openNode(this.tag, {
      sb: model.sb || 0,
      eb: model.eb || 0
    });
    if (model && model.hasOwnProperty('richText') && model.richText) {
      var r = this.map.r;
      model.richText.forEach(function (text) {
        r.render(xmlStream, text);
      });
    } else if (model) {
      this.map.t.render(xmlStream, model.text);
    }
    xmlStream.closeNode();
  },

  parseOpen: function parseOpen(node) {
    var name = node.name;
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (name === this.tag) {
      this.model = {
        sb: parseInt(node.attributes.sb, 10),
        eb: parseInt(node.attributes.eb, 10)
      };
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
            this.model.text = this.parser.model;
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
//# sourceMappingURL=phonetic-text-xform.js.map
