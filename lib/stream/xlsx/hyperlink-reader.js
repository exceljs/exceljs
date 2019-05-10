'use strict';

const events = require('events');
const Sax = require('sax');

const utils = require('../../utils/utils');
const Enums = require('../../doc/enums');
const RelType = require('../../xlsx/rel-type');

const HyperlinkReader = (module.exports = function(workbook, id) {
  // in a workbook, each sheet will have a number
  this.id = id;

  this._workbook = workbook;
});

utils.inherits(HyperlinkReader, events.EventEmitter, {
  get count() {
    return (this.hyperlinks && this.hyperlinks.length) || 0;
  },

  each(fn) {
    return this.hyperlinks.forEach(fn);
  },

  read(entry, options) {
    const self = this;
    let emitHyperlinks = false;
    let hyperlinks = null;
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

    const parser = Sax.createStream(true, {});
    parser.on('opentag', node => {
      if (node.name === 'Relationship') {
        const rId = node.attributes.Id;
        switch (node.attributes.Type) {
          case RelType.Hyperlink:
            const relationship = {
              type: Enums.RelationshipType.Styles,
              rId,
              target: node.attributes.Target,
              targetMode: node.attributes.TargetMode,
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
    parser.on('end', () => {
      self.emit('finished');
    });

    // create a down-stream flow-control to regulate the stream
    const flowControl = this._workbook.flowControl.createChild();
    flowControl.pipe(
      parser,
      { sync: true }
    );
    entry.pipe(flowControl);
  },
});
