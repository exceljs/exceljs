const XmlStream = require('../../../../utils/xml-stream');
const VmlShapeXform = require('./vml-shap-xform');
const BaseXform = require('../../base-xform');

class VmlDrawingXform extends BaseXform {
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
    xmlStream.openNode(this.tag, VmlDrawingXform.DRAWING_ATTRIBUTES);

    xmlStream.openNode('o:shapelayout', {'v:ext': 'edit'});
    xmlStream.leafNode('o:idmap', {'v:ext': 'edit', data: 1});
    xmlStream.closeNode();

    xmlStream.openNode('v:shapetype', {
      id: '_x0000_t202',
      coordsize: '21600,21600',
      'o:spt': 202,
      path: 'm,l,21600r21600,l21600,xe',
    });
    xmlStream.leafNode('v:stroke', {joinstyle: 'miter'});
    xmlStream.leafNode('v:path', {
      gradientshapeok: 't',
      'o:connecttype': 'rect',
    });
    xmlStream.closeNode();

    [...(model.comments || []), ...(model.hfImages || [])].forEach((item, index) => {
      this.map['v:shape'].render(xmlStream, item, index);
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
          hfImages: [],
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
        if (this.parser.model.type === 'hfimage') {
          this.model.hfImages.push(this.parser.model);
        } else {
          this.model.comments.push(this.parser.model);
        }
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
    model.hfImages.forEach(hfImage => {
      this.map['v:shape'].reconcile(hfImage, options);
    });
  }
}

VmlDrawingXform.DRAWING_ATTRIBUTES = {
  'xmlns:oa': 'urn:schemas-microsoft-com:office:activation',
  'xmlns:p': 'urn:schemas-microsoft-com:office:powerpoint',
  'xmlns:x': 'urn:schemas-microsoft-com:office:excel',
  'xmlns:o': 'urn:schemas-microsoft-com:office:office',
  'xmlns:v': 'urn:schemas-microsoft-com:vml',
};

module.exports = VmlDrawingXform;
