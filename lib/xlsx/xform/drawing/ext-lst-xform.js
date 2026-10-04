const BaseXform = require('../base-xform');

// a16:creationId identifies one drawing object. The last group is the
// picture sequence so two pictures do not share the same id.
function formatCreationId(sequence) {
  const n = Number(sequence) || 1;
  return `{00000000-0008-0000-0000-${n.toString(16).toUpperCase().padStart(12, '0')}}`;
}

class ExtLstXform extends BaseXform {
  get tag() {
    return 'a:extLst';
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag);
    xmlStream.openNode('a:ext', {
      uri: '{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}',
    });
    xmlStream.leafNode('a16:creationId', {
      'xmlns:a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
      id: (model && model.creationId) || formatCreationId(model && model.index),
    });
    xmlStream.closeNode();
    xmlStream.closeNode();
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        return true;
      default:
        return true;
    }
  }

  parseText() {}

  parseClose(name) {
    switch (name) {
      case this.tag:
        return false;
      default:
        // unprocessed internal nodes
        return true;
    }
  }
}

ExtLstXform.formatCreationId = formatCreationId;
module.exports = ExtLstXform;
