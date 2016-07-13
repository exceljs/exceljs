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
var TemplateXform = require('../template-xform');

// AppXform - used to render the app.xml
// Not used for parsing

var ContentTypesXform = module.exports = function() {
};


utils.inherits(ContentTypesXform, BaseXform, {
  STATIC_XFORMS: {
    doc: new TemplateXform(fs.readFileSync(__dirname + '/content-types.xml').toString())
  }
},{

  render: function(xmlStream, model) {
    ContentTypesXform.STATIC_XFORMS.doc.render(xmlStream, model);
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
