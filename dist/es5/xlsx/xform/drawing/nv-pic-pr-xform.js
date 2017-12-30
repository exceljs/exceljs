/**
 * Copyright (c) 2016-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');
var StaticXform = require('../static-xform');

var NvPicPrXform = module.exports = function () {};

utils.inherits(NvPicPrXform, BaseXform, {
  get tag() {
    return 'xdr:nvPicPr';
  },

  render: function render(xmlStream, model) {
    var xform = new StaticXform({
      tag: this.tag,
      c: [{
        tag: 'xdr:cNvPr', $: { id: model.index, name: 'Picture ' + model.index },
        c: [{
          tag: 'a:extLst',
          c: [{
            tag: 'a:ext', $: { uri: '{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}' },
            c: [{
              tag: 'a16:creationId',
              $: {
                'xmlns:a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
                id: '{00000000-0008-0000-0000-000002000000}'
              }
            }]
          }]
        }]
      }, {
        tag: 'xdr:cNvPicPr',
        c: [{ tag: 'a:picLocks', $: { noChangeAspect: '1' } }]
      }]
    });
    xform.render(xmlStream);
  }
});
//# sourceMappingURL=nv-pic-pr-xform.js.map
