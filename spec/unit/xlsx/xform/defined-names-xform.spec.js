'use strict';

var expect = require('chai').expect;

var _ = require('underscore');
var DefinedNamesXform = require('../../../../lib/xlsx/xform/defined-names-xform');

describe('DefinedNamesXform', function() {
  it('translate to xml', function() {
    var model = [
      {name: 'foo', ranges:['bar!$A1$C1']}
    ];
    var xml =
      '<definedNames>' +
        '<definedName name="foo">bar!$A1$C1</definedName>' +
      '</definedNames>';

    var dnx = new DefinedNamesXform(model);
      expect(dnx.xml).to.equal(xml);
  });

  it('parses ranges', function() {
    var dnx = new DefinedNamesXform();
    dnx.parseOpen({attributes: {name: 'single'}});
    dnx.parseText("'Sheet name'',$''!$'!$A$2");
    dnx.parseClose();

    dnx.parseOpen({attributes: {name: 'range'}});
    dnx.parseText("'Sheet name'',$''!$'!$A$1:$B$1");
    dnx.parseClose();

    dnx.parseOpen({attributes: {name: 'multiple'}});
    dnx.parseText("'Sheet name''$''!$'!$A$2,'Sheet name''$''!$'!$H$7,'Sheet3!'!$A$1,Sheet3!$A$1");
    dnx.parseClose();

    expect(dnx.model).to.deep.equal([
      {name: 'single', ranges: ["'Sheet name'',$''!$'!$A$2"]},
      {name: 'range', ranges: ["'Sheet name'',$''!$'!$A$1:$B$1"]},
      {
        name: 'multiple',
        ranges: [
          "'Sheet name''$''!$'!$A$2",
          "'Sheet name''$''!$'!$H$7",
          "'Sheet3!'!$A$1",
          "Sheet3!$A$1"
        ]
      }
    ]);
  });

});