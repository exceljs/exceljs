/**
 * Copyright (c) 2015 Guyon Roche
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

// used for rendering the [Content_Types].xml file
// not used for parsing
var ContentTypesXform = module.exports = function() {
};


utils.inherits(ContentTypesXform, BaseXform, {
  PROPERTY_ATTRIBUTES: {
    xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types'
  }
},{

  render: function(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);

    xmlStream.openNode('Types', ContentTypesXform.PROPERTY_ATTRIBUTES);

    xmlStream.leafNode('Default', {Extension:'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml'});
    xmlStream.leafNode('Default', {Extension:'xml', ContentType: 'application/xml'});

    xmlStream.leafNode('Override', {PartName:'/xl/workbook.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'});

    model.worksheets.forEach(function(worksheet) {
      var name = '/xl/worksheets/sheet' + worksheet.id + '.xml';
      xmlStream.leafNode('Override', {PartName:name, ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'});
    });

    xmlStream.leafNode('Override', {PartName:'/xl/theme/theme1.xml', ContentType: 'application/vnd.openxmlformats-officedocument.theme+xml'});
    xmlStream.leafNode('Override', {PartName:'/xl/styles.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'});
    xmlStream.leafNode('Override', {PartName:'/xl/sharedStrings.xml', ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'});
    xmlStream.leafNode('Override', {PartName:'/docProps/core.xml', ContentType: 'application/vnd.openxmlformats-package.core-properties+xml'});
    xmlStream.leafNode('Override', {PartName:'/docProps/app.xml', ContentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml'});

    xmlStream.closeNode();
  },

  parseOpen: function() {
    return false;
  },
  parseText: function() {
  },
  parseClose: function() {
    return false;
  }
});
