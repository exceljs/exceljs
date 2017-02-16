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

var props = {
  creator: 'Author',
  title: 'Title',
  subject: 'Subject',
  description: 'Comments',
  language: 'Language',
  keywords: 'Tags',
  category: 'Categories',
  identifier: 'Identifier'
};

var CoreXform = module.exports = function() {
  this.map = {
    'dc:creator': new StringXform({tag: 'dc:creator'}),
    'dc:title': new StringXform({tag: 'dc:title'}),
    'dc:subject': new StringXform({tag: 'dc:subject'}),
    'dc:description': new StringXform({tag: 'dc:description'}),
    'dc:identifier': new StringXform({tag: 'dc:identifier'}),
    'dc:language': new StringXform({tag: 'dc:language'}),
    'cp:keywords': new StringXform({tag: 'cp:keywords'}),
    'cp:category': new StringXform({tag: 'cp:category'}),
    'cp:lastModifiedBy': new StringXform({tag: 'cp:lastModifiedBy'}),
    'cp:lastPrinted': new DateXform({tag: 'cp:lastPrinted', format: CoreXform.DateFormat}),
    'cp:revision': new DateXform({tag: 'cp:revision'}),
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
    this.map['dc:title'].render(xmlStream, model.title);
    this.map['dc:subject'].render(xmlStream, model.subject);
    this.map['dc:description'].render(xmlStream, model.description);
    this.map['dc:identifier'].render(xmlStream, model.identifier);
    this.map['dc:language'].render(xmlStream, model.language);
    this.map['cp:keywords'].render(xmlStream, model.keywords);
    this.map['cp:category'].render(xmlStream, model.category);
    this.map['cp:lastModifiedBy'].render(xmlStream, model.lastModifiedBy);
    this.map['cp:lastPrinted'].render(xmlStream, model.lastPrinted);
    this.map['cp:revision'].render(xmlStream, model.revision);
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
            title: this.map['dc:title'].model,
            subject: this.map['dc:subject'].model,
            description: this.map['dc:description'].model,
            identifier: this.map['dc:identifier'].model,
            language: this.map['dc:language'].model,
            keywords: this.map['cp:keywords'].model,
            category: this.map['cp:category'].model,
            lastModifiedBy: this.map['cp:lastModifiedBy'].model,
            lastPrinted: this.map['cp:lastPrinted'].model,
            revision: this.map['cp:revision'].model,
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
