const BaseXform = require('../base-xform');

const VmlAnchorXform = require('./vml-anchor-xform');
const VmlTextboxXform = require('./vml-textbox-xform');

// render the triangle in the cell for the comment
class VmlNoteXform extends BaseXform {
  constructor() {
    super();
    this.map = {
      'x:Anchor': new VmlAnchorXform(),
      'v:textbox': new VmlTextboxXform(),
    };
  }

  get tag() {
    return 'v:shape';
  }

  normalizeAttributes(model, index) {
    return {
      id: `_x0000_s${1025 + index}`,
      type: '#_x0000_t202',
      style: 'position:absolute; margin-left:105.3pt;margin-top:10.5pt;width:97.8pt;height:59.1pt;z-index:1;visibility:hidden',
      fillcolor: 'infoBackground [80]',
      strokecolor: 'none [81]',
      'o:insetmode': model.note && model.note.insetmode,
    };
  }

  render(xmlStream, model, index) {
    xmlStream.openNode('v:shape', this.normalizeAttributes(model, index));

    xmlStream.leafNode('v:fill', {color2: 'infoBackground [80]'});
    xmlStream.leafNode('v:shadow', {color: 'none [81]', obscured: 't'});
    xmlStream.leafNode('v:path', {'o:connecttype': 'none'});
    this.map['v:textbox'].render(xmlStream, model);

    xmlStream.openNode('x:ClientData', {ObjectType: 'Note'});
    xmlStream.leafNode('x:MoveWithCells');
    xmlStream.leafNode('x:SizeWithCells');

    this.map['x:Anchor'].render(xmlStream, model);

    xmlStream.leafNode('x:AutoFill', null, 'False');
    xmlStream.leafNode('x:Row', null, model.refAddress.row - 1);
    xmlStream.leafNode('x:Column', null, model.refAddress.col - 1);
    xmlStream.closeNode();

    xmlStream.closeNode();
  }

  parseOpen(node) {
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          anchor: [],
          note: {
            insetmode: node.attributes['o:insetmode'],
          },
        };
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
  }

  parseText() {}

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.model.anchor = this.map['x:Anchor'].model;
        this.model.note = Object.assign(this.model.note, this.map['v:textbox'].model);
        return false;
      default:
        return true;
    }
  }
}

module.exports = VmlNoteXform;
