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
var StringXform = require('../simple/string-xform');

var AppHeadingPairsXform = require('./app-heading-pairs-xform');
var AppTitleOfPartsXform = require('./app-titles-of-parts-xform');

var props = {
  company: 'Company',
  manager: 'Manager',
  worksheets: [{name: 'sheet1'}]
};

var AppXform = module.exports = function() {
  this.map = {
    'Company': new StringXform({tag: 'Company'}),
    'Manager': new StringXform({tag: 'Manager'}),
    'HeadingPairs': new AppHeadingPairsXform(),
    'TitleOfParts': new AppTitleOfPartsXform()
  }
};

AppXform.DateFormat = function(dt) {
  return dt.toISOString().replace(/\.\d{3}/,'');
};
AppXform.DateAttrs = {'xsi:type': 'dcterms:W3CDTF'};

AppXform.PROPERTY_ATTRIBUTES = {
  'xmlns': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
  'xmlns:vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
};

utils.inherits(AppXform, BaseXform, {
  render: function(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);

    xmlStream.openNode('Properties', AppXform.PROPERTY_ATTRIBUTES);
    
    xmlStream.leafNode('Application',  undefined, 'Microsoft Excel');
    xmlStream.leafNode('DocSecurity',  undefined, '0');
    xmlStream.leafNode('ScaleCrop',  undefined, 'false');

    this.map['HeadingPairs'].render(xmlStream, model.worksheets);
    this.map['TitleOfParts'].render(xmlStream, model.worksheets);
    this.map['Company'].render(xmlStream, model.company || '');
    this.map['Manager'].render(xmlStream, model.manager);

    xmlStream.leafNode('LinksUpToDate',  undefined, 'false');
    xmlStream.leafNode('SharedDoc',  undefined, 'false');
    xmlStream.leafNode('HyperlinksChanged',  undefined, 'false');
    xmlStream.leafNode('AppVersion',  undefined, '16.0300');

    xmlStream.closeNode();
  },

  parseOpen: function(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    } else {
      switch(node.name) {
        case 'Properties':
          return true;
        default:
          this.parser = this.map[node.name];
          if (this.parser) {
            this.parser.parseOpen(node);
            return true;
          }
          
          // there's a lot we don't bother to parse
          return false;
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
        case 'Properties':
          this.model = {
            worksheets: this.map['TitleOfParts'].model,
            company: this.map['Company'].model,
            manager: this.map['Manager'].model
          };
          return false;
        default:
          return true;
      }
    }
  }
});
