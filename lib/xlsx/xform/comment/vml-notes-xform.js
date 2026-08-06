const XmlStream = require('../../../utils/xml-stream');

const BaseXform = require('../base-xform');
const VmlShapeXform = require('./vml-shape-xform');

// This class is (currently) single purposed to insert the triangle
// drawing icons on commented cells
class VmlNotesXform extends BaseXform {
  constructor() {
    super();
    this.map = {
      'v:shape': new VmlShapeXform(),
    };
  }

  get tag() {
    return 'xml';
  }

  render(xmlStream, model) {
    xmlStream.openXml(XmlStream.StdDocAttributes);
    xmlStream.openNode(this.tag, VmlNotesXform.DRAWING_ATTRIBUTES);

    xmlStream.openNode('o:shapelayout', {'v:ext': 'edit'});
    // Excel allocates VML shape ids workbook-globally in blocks of 1024 per drawing part and
    // expects each part to reserve a distinct block (ISO 29500 §14.2.2.14 "idmap"). Emitting
    // block "1" and shape ids _x0000_s1025... for every sheet makes Excel report workbooks with
    // comments on two or more sheets as corrupt ("We found a problem with some content...").
    // Use the sheet id as block index (sheet 1 -> data="1", ids 1025+; sheet 2 -> data="2",
    // ids 2049+; ...), which is how Excel itself numbers these parts.
    xmlStream.leafNode('o:idmap', {'v:ext': 'edit', data: model.id || 1});
    xmlStream.closeNode();

    xmlStream.openNode('v:shapetype', {
      id: '_x0000_t202',
      coordsize: '21600,21600',
      'o:spt': 202,
      path: 'm,l,21600r21600,l21600,xe',
    });
    xmlStream.leafNode('v:stroke', {joinstyle: 'miter'});
    xmlStream.leafNode('v:path', {gradientshapeok: 't', 'o:connecttype': 'rect'});
    xmlStream.closeNode();

    model.comments.forEach((item, index) => {
      // Offset shape ids into this sheet's 1024-id block so they are unique workbook-wide.
      this.map['v:shape'].render(xmlStream, item, (((model.id || 1) - 1) * 1024) + index);
    });

    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          comments: [],
        };
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  parseText(text) {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  parseClose(name) {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model.comments.push(this.parser.model);
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        return false;
      default:
        // could be some unrecognised tags
        return true;
    }
  }

  reconcile(model, options) {
    model.anchors.forEach(anchor => {
      if (anchor.br) {
        this.map['xdr:twoCellAnchor'].reconcile(anchor, options);
      } else {
        this.map['xdr:oneCellAnchor'].reconcile(anchor, options);
      }
    });
  }
}

VmlNotesXform.DRAWING_ATTRIBUTES = {
  'xmlns:v': 'urn:schemas-microsoft-com:vml',
  'xmlns:o': 'urn:schemas-microsoft-com:office:office',
  'xmlns:x': 'urn:schemas-microsoft-com:office:excel',
};

module.exports = VmlNotesXform;
