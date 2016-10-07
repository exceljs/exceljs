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

var utils = require('../../../utils/utils');
var XmlStream = require('../../../utils/xml-stream');
var BaseXform = require('../base-xform');
var SharedStringXform = require('./shared-string-xform');

var SharedStringsXform = module.exports = function(model) {
  this.model = model || {
    values: [],
    count: 0
  };
  this.hash = {};
  this.rich = {};
};

utils.inherits(SharedStringsXform, BaseXform, {
  get sharedStringXform() { return this._sharedStringXform || (this._sharedStringXform = new SharedStringXform()); },

  get values() {
    return this.model.values;
  },
  get uniqueCount() {
    return this.model.values.length;
  },
  get count() {
    return this.model.count;
  },

  getString: function(index) {
    return this.model.values[index];
  },

  add: function(value) {
    return value.richText ?
      this.addRichText(value) :
      this.addText(value);
  },
  addText: function(value) {
    var index = this.hash[value];
    if (index === undefined) {
      index = this.hash[value] = this.model.values.length;
      this.model.values.push(value);
    }
    this.model.count++;
    return index;
  },
  addRichText: function(value) {
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

  render: function(xmlStream, model) {
    model = model || this._values;
    xmlStream.openXml(XmlStream.StdDocAttributes);

    xmlStream.openNode('sst', {
      xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      count: model.count,
      uniqueCount: model.values.length
    });

    var sx =  this.sharedStringXform;
    model.values.forEach(function(sharedString) {
      sx.render(xmlStream, sharedString);
    });
    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      switch(node.name) {
        case 'sst':
          return true;
        case 'si':
          this.parser = this.sharedStringXform;
          this.parser.parseOpen(node);
          return true;
        default:
          throw new Error('Unexpected xml node in parseOpen: ' + JSON.stringify(node));
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
        this.model.values.push(this.parser.model);
        this.model.count++;
        this.parser = undefined;
      }
      return true;
    } else {
      switch(name) {
        case 'sst':
          return false;
        default:
          throw new Error('Unexpected xml node in parseClose: ' + name);
      }
    }
  }
});
