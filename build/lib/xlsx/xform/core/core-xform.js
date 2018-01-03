/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var XmlStream = require('../../../utils/xml-stream');
var BaseXform = require('../base-xform');
var DateXform = require('../simple/date-xform');
var StringXform = require('../simple/string-xform');

var CoreXform = module.exports = function () {
  this.map = {
    'creator': new StringXform({ tag: 'creator', ns: 'dc' }),
    'title': new StringXform({ tag: 'title', ns: 'dc' }),
    'subject': new StringXform({ tag: 'subject', ns: 'dc' }),
    'description': new StringXform({ tag: 'description', ns: 'dc' }),
    'identifier': new StringXform({ tag: 'identifier', ns: 'dc' }),
    'language': new StringXform({ tag: 'language', ns: 'dc' }),
    'keywords': new StringXform({ tag: 'keywords', ns: 'cp' }),
    'category': new StringXform({ tag: 'category', ns: 'cp' }),
    'lastModifiedBy': new StringXform({ tag: 'lastModifiedBy', ns: 'cp' }),
    'lastPrinted': new DateXform({ tag: 'lastPrinted', ns: 'cp', format: CoreXform.DateFormat }),
    'revision': new DateXform({ tag: 'revision', ns: 'cp' }),
    'version': new StringXform({ tag: 'version', ns: 'cp' }),
    'contentStatus': new StringXform({ tag: 'contentStatus', ns: 'cp' }),
    'created': new DateXform({ tag: 'created', ns: 'dcterms', attrs: CoreXform.DateAttrs, format: CoreXform.DateFormat }),
    'modified': new DateXform({ tag: 'modified', ns: 'dcterms', attrs: CoreXform.DateAttrs, format: CoreXform.DateFormat })
  };
};

CoreXform.DateFormat = function (dt) {
  return dt.toISOString().replace(/[.]\d{3}/, '');
};
CoreXform.DateAttrs = { 'xsi:type': 'dcterms:W3CDTF' };

CoreXform.CORE_PROPERTY_ATTRIBUTES = {
  'xmlns:cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
  'xmlns:dcterms': 'http://purl.org/dc/terms/',
  'xmlns:dcmitype': 'http://purl.org/dc/dcmitype/',
  'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance'
};

utils.inherits(CoreXform, BaseXform, {
  render: function render(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);

    xmlStream.openNode('cp:coreProperties', CoreXform.CORE_PROPERTY_ATTRIBUTES);

    this.map['creator'].render(xmlStream, model.creator);
    this.map['title'].render(xmlStream, model.title);
    this.map['subject'].render(xmlStream, model.subject);
    this.map['description'].render(xmlStream, model.description);
    this.map['identifier'].render(xmlStream, model.identifier);
    this.map['language'].render(xmlStream, model.language);
    this.map['keywords'].render(xmlStream, model.keywords);
    this.map['category'].render(xmlStream, model.category);
    this.map['lastModifiedBy'].render(xmlStream, model.lastModifiedBy);
    this.map['lastPrinted'].render(xmlStream, model.lastPrinted);
    this.map['revision'].render(xmlStream, model.revision);
    this.map['version'].render(xmlStream, model.version);
    this.map['contentStatus'].render(xmlStream, model.contentStatus);
    this.map['created'].render(xmlStream, model.created);
    this.map['modified'].render(xmlStream, model.modified);

    xmlStream.closeNode();
  },

  parseOpen: function parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'coreProperties':
        return true;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
          return true;
        }
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
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case 'coreProperties':
        this.model = {
          creator: this.map['creator'].model,
          title: this.map['title'].model,
          subject: this.map['subject'].model,
          description: this.map['description'].model,
          identifier: this.map['identifier'].model,
          language: this.map['language'].model,
          keywords: this.map['keywords'].model,
          category: this.map['category'].model,
          lastModifiedBy: this.map['lastModifiedBy'].model,
          lastPrinted: this.map['lastPrinted'].model,
          revision: this.map['revision'].model,
          contentStatus: this.map['contentStatus'].model,
          created: this.map['created'].model,
          modified: this.map['modified'].model
        };
        return false;
      default:
        throw new Error('Unexpected xml node in parseClose: ' + name);
    }
  }
});
//# sourceMappingURL=core-xform.js.map
