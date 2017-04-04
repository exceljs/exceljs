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
var BaseXform = require('../base-xform');
var StaticXform = require('../static-xform');

var NvPicPrXform = module.exports = function() {
};

utils.inherits(NvPicPrXform, BaseXform, {
  get tag() { return 'xdr:nvPicPr'; },

  render: function(xmlStream, model) {
    var xform = new StaticXform({
      tag: this.tag,
      c: [
        {
          tag: 'xdr:cNvPr', $: {id: model.index, name: 'Picture ' + model.index},
          c: [{
            tag: 'a:extLst',
            c: [{
              tag: 'a:ext', $: {uri: '{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}'},
              c: [{
                tag: 'a16:creationId',
                $: {
                  'xmlns:a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
                  id: '{00000000-0008-0000-0000-000002000000}',
                }
              }]
            }]
          }]
        },
        {
          tag: 'xdr:cNvPicPr',
          c: [{tag: 'a:picLocks', $: {noChangeAspect: '1'}}]
        }
      ]
    });
    xform.render(xmlStream);
  },
});
