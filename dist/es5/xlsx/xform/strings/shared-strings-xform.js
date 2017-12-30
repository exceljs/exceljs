/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var XmlStream = require('../../../utils/xml-stream');
var BaseXform = require('../base-xform');
var SharedStringXform = require('./shared-string-xform');

var SharedStringsXform = module.exports = function (model) {
  this.model = model || {
    values: [],
    count: 0
  };
  this.hash = {};
  this.rich = {};
};

utils.inherits(SharedStringsXform, BaseXform, {
  get sharedStringXform() {
    return this._sharedStringXform || (this._sharedStringXform = new SharedStringXform());
  },

  get values() {
    return this.model.values;
  },
  get uniqueCount() {
    return this.model.values.length;
  },
  get count() {
    return this.model.count;
  },

  getString: function getString(index) {
    return this.model.values[index];
  },

  add: function add(value) {
    return value.richText ? this.addRichText(value) : this.addText(value);
  },
  addText: function addText(value) {
    var index = this.hash[value];
    if (index === undefined) {
      index = this.hash[value] = this.model.values.length;
      this.model.values.push(value);
    }
    this.model.count++;
    return index;
  },
  addRichText: function addRichText(value) {
    // TODO: add WeakMap here
    var xml = this.sharedStringXform.toXml(value);
    var index = this.rich[xml];
    if (index === undefined) {
      index = this.rich[xml] = this.model.values.length;
      this.model.values.push(value);
    }
    this.model.count++;
    return index;
  },

  // <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  // <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="<%=totalRefs%>" uniqueCount="<%=count%>">
  //   <si><t><%=text%></t></si>
  //   <si><r><rPr></rPr><t></t></r></si>
  // </sst>

  render: function render(xmlStream, model) {
    model = model || this._values;
    xmlStream.openXml(XmlStream.StdDocAttributes);

    xmlStream.openNode('sst', {
      xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      count: model.count,
      uniqueCount: model.values.length
    });

    var sx = this.sharedStringXform;
    model.values.forEach(function (sharedString) {
      sx.render(xmlStream, sharedString);
    });
    xmlStream.closeNode();
  },

  parseOpen: function parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'sst':
        return true;
      case 'si':
        this.parser = this.sharedStringXform;
        this.parser.parseOpen(node);
        return true;
      default:
        throw new Error('Unexpected xml node in parseOpen: ' + JSON.stringify(node));
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
        this.model.values.push(this.parser.model);
        this.model.count++;
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case 'sst':
        return false;
      default:
        throw new Error('Unexpected xml node in parseClose: ' + name);
    }
  }
});
//# sourceMappingURL=shared-strings-xform.js.map
