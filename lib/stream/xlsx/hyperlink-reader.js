const {EventEmitter} = require('events');
const SAXStream = require('../../utils/sax-stream');

const Enums = require('../../doc/enums');
const RelType = require('../../xlsx/rel-type');

class HyperlinkReader extends EventEmitter {
  constructor(workbook, id) {
    super();

    // in a workbook, each sheet will have a number
    this.id = id;

    this._workbook = workbook;
  }

  get count() {
    return (this.hyperlinks && this.hyperlinks.length) || 0;
  }

  each(fn) {
    return this.hyperlinks.forEach(fn);
  }

  read(entry, options) {
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
      this.emit('finished');
      return;
    }

    const saxStream = new SAXStream(['opentag']);
    saxStream.on('data', events => {
      for (const {eventType, value} of events) {
        if (eventType === 'opentag') {
          const node = value;
          if (node.name === 'Relationship') {
            const rId = node.attributes.Id;
            switch (node.attributes.Type) {
              case RelType.Hyperlink:
                {
                  const relationship = {
                    type: Enums.RelationshipType.Styles,
                    rId,
                    target: node.attributes.Target,
                    targetMode: node.attributes.TargetMode,
                  };
                  if (emitHyperlinks) {
                    this.emit('hyperlink', relationship);
                  } else {
                    hyperlinks[relationship.rId] = relationship;
                  }
                }
                break;

              default:
                break;
            }
          }
        }
      }
    });
    saxStream.on('end', () => {
      this.emit('finished');
    });
    entry.pipe(saxStream);
  }
}

module.exports = HyperlinkReader;
