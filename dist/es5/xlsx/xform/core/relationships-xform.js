/**
 * Copyright (c) 2016 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var utils = require('../../../utils/utils');
var XmlStream = require('../../../utils/xml-stream');
var BaseXform = require('../base-xform');

var RelationshipXform = require('./relationship-xform');

var RelationshipsXform = module.exports = function () {
  this.map = {
    Relationship: new RelationshipXform()
  };
};

utils.inherits(RelationshipsXform, BaseXform, {
  RELATIONSHIPS_ATTRIBUTES: { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' }
}, {
  render: function render(xmlStream, model) {
    model = model || this._values;
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode('Relationships', RelationshipsXform.RELATIONSHIPS_ATTRIBUTES);

    var self = this;
    model.forEach(function (relationship) {
      self.map.Relationship.render(xmlStream, relationship);
    });

    xmlStream.closeNode();
  },

  parseOpen: function parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'Relationships':
        this.model = [];
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
        this.model.push(this.parser.model);
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case 'Relationships':
        return false;
      default:
        throw new Error('Unexpected xml node in parseClose: ' + name);
    }
  }
});
//# sourceMappingURL=relationships-xform.js.map
