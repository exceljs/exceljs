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
var DateXform = require('../simple/date-xform');
var StringXform = require('../simple/string-xform');

// <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
// <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
//   <dc:creator><%=creator%></dc:creator>
// <cp:lastModifiedBy><%=lastModifiedBy%></cp:lastModifiedBy>
// <dcterms:created xsi:type="dcterms:W3CDTF"><%=created.toISOString().replace(/\.\d{3}/,"")%></dcterms:created>
// <dcterms:modified xsi:type="dcterms:W3CDTF"><%=modified.toISOString().replace(/\.\d{3}/,"")%></dcterms:modified>
// </cp:coreProperties>

var CoreXform = module.exports = function() {
  this.map = {
    'dc:creator': new StringXform({tag: 'dc:creator'}),
    'cp:lastModifiedBy': new StringXform({tag: 'cp:lastModifiedBy'}),
    'dcterms:created': new DateXform({tag: 'dcterms:created', attrs: CoreXform.DateAttrs, format: CoreXform.DateFormat}),
    'dcterms:modified': new DateXform({tag: 'dcterms:modified', attrs: CoreXform.DateAttrs, format: CoreXform.DateFormat})
  }
};

CoreXform.DateFormat = function(dt) {
  return dt.toISOString().replace(/\.\d{3}/,'');
};
CoreXform.DateAttrs = {'xsi:type': 'dcterms:W3CDTF'};

CoreXform.CORE_PROPERTY_ATTRIBUTES = {
  'xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
  'xmlns:dcterms': 'http://purl.org/dc/terms/',
  'xmlns:dcmitype': 'http://purl.org/dc/dcmitype/',
  'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance'
};

utils.inherits(CoreXform, BaseXform, {
  render: function(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);

    xmlStream.openNode('cp:coreProperties', CoreXform.CORE_PROPERTY_ATTRIBUTES);

    this.map['dc:creator'].render(xmlStream, model.creator);
    this.map['cp:lastModifiedBy'].render(xmlStream, model.lastModifiedBy);
    this.map['dcterms:created'].render(xmlStream, model.created);
    this.map['dcterms:modified'].render(xmlStream, model.modified);

    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      switch(node.name) {
        case 'cp:coreProperties':
          return true;
        default:
          this.parser = this.map[node.name];
          if (this.parser) {
            this.parser.parseOpen(node);
            return true;
          }
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
        this.parser = undefined;
      }
      return true;
    } else {
      switch(name) {
        case 'cp:coreProperties':
          this.model = {
            creator: this.map['dc:creator'].model,
            lastModifiedBy: this.map['cp:lastModifiedBy'].model,
            created: this.map['dcterms:created'].model,
            modified: this.map['dcterms:modified'].model
          };
          return false;
        default:
          throw new Error('Unexpected xml node in parseClose: ' + name);
      }
    }
  }
});
