/**
 * Copyright (c) 2015-2017 Guyon Roche
 * LICENCE: MIT - please refer to LICENCE file included with this module
 * or https://github.com/guyonroche/exceljs/blob/master/LICENSE
 */

'use strict';

var events = require('events');
var Sax = require('sax');

var utils = require('../../utils/utils');
var Enums = require('../../doc/enums');
var RelType = require('../../xlsx/rel-type');

var HyperlinkReader = module.exports = function (workbook, id) {
  // in a workbook, each sheet will have a number
  this.id = id;

  this._workbook = workbook;
};

utils.inherits(HyperlinkReader, events.EventEmitter, {

  get count() {
    return this.hyperlinks && this.hyperlinks.length || 0;
  },

  each: function each(fn) {
    return this.hyperlinks.forEach(fn);
  },

  read: function read(entry, options) {
    var self = this;
    var emitHyperlinks = false;
    var hyperlinks = null;
    switch (options.hyperlinks) {
      case 'emit':
        emitHyperlinks = true;
        break;
      case 'cache':
        this.hyperlinks = hyperlinks = {};
        break;
      default:
        break;
    }
    if (!emitHyperlinks && !hyperlinks) {
      entry.autodrain();
      self.emit('finished');
      return;
    }

    var parser = Sax.createStream(true, {});
    parser.on('opentag', function (node) {
      if (node.name === 'Relationship') {
        var rId = node.attributes.Id;
        switch (node.attributes.Type) {
          case RelType.Hyperlink:
            var relationship = {
              type: Enums.RelationshipType.Styles,
              rId: rId,
              target: node.attributes.Target,
              targetMode: node.attributes.TargetMode
            };
            if (emitHyperlinks) {
              self.emit('hyperlink', relationship);
            } else {
              hyperlinks[relationship.rId] = relationship;
            }
            break;
          default:
            break;
        }
      }
    });
    parser.on('end', function () {
      self.emit('finished');
    });

    // create a down-stream flow-control to regulate the stream
    var flowControl = this._workbook.flowControl.createChild();
    flowControl.pipe(parser, { sync: true });
    entry.pipe(flowControl);
  }
});
//# sourceMappingURL=hyperlink-reader.js.map
