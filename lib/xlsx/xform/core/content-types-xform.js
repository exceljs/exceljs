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

var fs  = require('fs');
var events = require('events');

var utils = require('../../../utils/utils');

var BaseXform = require('../base-xform');
var XmlStream = require('../../../utils/xml-stream');

var ContentTypesXform = module.exports = function() {
};

utils.inherits(ContentTypesXform, BaseXform, {
  TYPES_PROPERTY_ATTRIBUTES: {
    xmlns:'http://schemas.openxmlformats.org/package/2006/content-types'
  },
  STD_DEFAULTS: [
    { Extension:'rels', ContentType:'application/vnd.openxmlformats-package.relationships+xml'},
    { Extension:'xml', ContentType:'application/xml'},
  ],
  STD_OVERRIDES: [
    { PartName:'/xl/workbook.xml', ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' },
    { PartName:'/xl/theme/theme1.xml', ContentType:'application/vnd.openxmlformats-officedocument.theme+xml' },
    { PartName:'/xl/styles.xml', ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml' },
    { PartName:'/xl/sharedStrings.xml', ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml' },
    { PartName:'/docProps/core.xml', ContentType:'application/vnd.openxmlformats-package.core-properties+xml' },
    { PartName:'/docProps/app.xml', ContentType:'application/vnd.openxmlformats-officedocument.extended-properties+xml' },
  ],
},{

  render: function(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);

    xmlStream.openNode('Types', ContentTypesXform.TYPES_PROPERTY_ATTRIBUTES);

    var mediaHash =  {};
    (model.media || []).forEach(function(medium) {
      if (medium.type === 'image') {
        var imageType = medium.image.type;
        if (!mediaHash[imageType]) {
          mediaHash[imageType] = true;
          xmlStream.leafNode('Default', {Extension: imageType, ContentType: 'image/' + imageType});
        }
      }
    });

    ContentTypesXform.STD_DEFAULTS.forEach(function(value) {
      xmlStream.leafNode('Default', value);
    });

    model.worksheets.forEach(function(sheet) {
      xmlStream.leafNode('Override', {
        PartName: '/xl/worksheets/sheet' + sheet.id + '.xml',
        ContentType:  'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
      });
    });

    ContentTypesXform.STD_OVERRIDES.forEach(function(value) {
      xmlStream.leafNode('Override', value);
    });

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
