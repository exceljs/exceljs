const BaseXform = require('../base-xform');

const VmlAnchorXform = require('./vml-anchor-xform');

// render the triangle in the cell for the comment
class VmlNoteXform extends BaseXform {
  get tag() {
    return 'v:shape';
  }

  render(xmlStream, model, index) {
    xmlStream.openNode('v:shape', VmlNoteXform.V_SHAPE_ATTRIBUTES(index));

    xmlStream.leafNode('v:fill', {color2: 'infoBackground [80]'});
    xmlStream.leafNode('v:shadow', {color: 'none [81]', obscured: 't'});
    xmlStream.leafNode('v:path', {'o:connecttype': 'none'});

    xmlStream.openNode('v:textbox', {style: 'mso-direction-alt:auto'});
    xmlStream.leafNode('div', {style: 'text-align:left'});
    xmlStream.closeNode();

    xmlStream.openNode('x:ClientData', {ObjectType: 'Note'});
    xmlStream.leafNode('x:MoveWithCells');
    xmlStream.leafNode('x:SizeWithCells');

    VmlNoteXform.vmlAnchorXform.render(xmlStream, model);

    xmlStream.leafNode('x:AutoFill', null, 'False');
    xmlStream.leafNode('x:Row', null, model.refAddress.row - 1);
    xmlStream.leafNode('x:Column', null, model.refAddress.col - 1);
    xmlStream.closeNode();

    xmlStream.closeNode();
  }
}

module.exports = VmlNoteXform;

VmlNoteXform.V_SHAPE_ATTRIBUTES = index => ({
  id: `_x0000_s${1025 + index}`,
  type: '#_x0000_t202',
  style: 'position:absolute; margin-left:105.3pt;margin-top:10.5pt;width:97.8pt;height:59.1pt;z-index:1;visibility:hidden',
  fillcolor: 'infoBackground [80]',
  strokecolor: 'none [81]',
  'o:insetmode': 'auto',
});

VmlNoteXform.vmlAnchorXform = new VmlAnchorXform();