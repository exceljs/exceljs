const BaseXform = require('../../base-xform');
const VmlTextboxXform = require('../../comment/vml-textbox-xform');
const VmlClientDataXform = require('../../comment/vml-client-data-xform');
const VmlPathform = require('./vml-path-xform');
const VmlFillform = require('./vml-fill-xform');
const VmlStrokeform = require('./vml-stroke-xform');
const VmlImageform = require('./vml-image.xform');
const OLockform = require('./o-lock-xform');
const VmlShadowform = require('./vml-shadow-xform');

const VmlShapeType = {
  commont: 'commont',
  hfimage: 'hfimage',
};

class VShapeXform extends BaseXform {
  constructor() {
    super();
    this.map = {
      'v:textbox': new VmlTextboxXform(),
      'x:ClientData': new VmlClientDataXform(),
      'v:path': new VmlPathform(),
      'v:fill': new VmlFillform(),
      'v:stroke': new VmlStrokeform(),
      'v:imagedata': new VmlImageform(),
      'v:shadow': new VmlShadowform(),
      'o:lock': new OLockform(),
    };
  }

  get tag() {
    return 'v:shape';
  }

  render(xmlStream, model, index) {
    xmlStream.openNode('v:shape', VShapeXform.V_SHAPE_ATTRIBUTES(model, index));
    if (model.shadow) {
      this.map['v:shadow'].render(xmlStream, model.shadow);
    }
    if (model.path) {
      this.map['v:path'].render(xmlStream, model.path);
    }
    if (model.fill) {
      this.map['v:fill'].render(xmlStream, model.fill);
    }
    if (model.stroke) {
      this.map['v:stroke'].render(xmlStream, model.stroke);
    }
    if (model.imagedata) {
      this.map['v:imagedata'].render(xmlStream, model.imagedata);
    }
    if (model.stroke) {
      this.map['o:lock'].render(xmlStream, model.stroke);
    }

    if (model.margins && model.margins.inset) {
      this.map['v:textbox'].render(xmlStream, model);
    }

    this.map['x:ClientData'].render(xmlStream, model);

    xmlStream.closeNode();
  }

  parseOpen(node) {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }

    const isHF = node.attributes.type === '#_x0000_t75';
    switch (node.name) {
      case this.tag:
        this.reset();
        if (isHF) {
          const styleObj = node.attributes.style.split(';').reduce((obj, str) => {
            if (str === '') {
              return obj;
            }
            const [key, value] = str.split(':');
            obj[key] = value;
            return obj;
          }, {});
          this.model = {
            type: VmlShapeType.hfimage,
            id: node.attributes.id,
            style: styleObj,
          };
        } else {
          this.model = {
            type: VmlShapeType.commont,
            margins: {
              insetmode: node.attributes['o:insetmode'],
            },
            anchor: '',
            editAs: '',
            protection: {},
          };
        }
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
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        if (this.model.margins) {
          this.model.margins.inset =
            this.map['v:textbox'].model && this.map['v:textbox'].model.inset;
        }
        this.model.protection =
          this.map['x:ClientData'].model && this.map['x:ClientData'].model.protection;
        this.model.anchor = this.map['x:ClientData'].model && this.map['x:ClientData'].model.anchor;
        this.model.editAs = this.map['x:ClientData'].model && this.map['x:ClientData'].model.editAs;
        this.model.path = this.map['v:path'].model;
        this.model.fill = this.map['v:fill'].model;
        this.model.stroke = this.map['v:stroke'].model;
        this.model.imagedata = this.map['v:imagedata'].model;
        this.model.shadow = this.map['v:shadow'].model;

        return false;
      default:
        return true;
    }
  }

  reconcile(model, options) {
    this.map['v:imagedata'].reconcile(model.imagedata, options);
  }
}

VShapeXform.V_SHAPE_ATTRIBUTES = (model, index) => {
  if (model.type === VmlShapeType.hfimage) {
    return {
      id: model.id,
      style:
        'position:absolute;left:0pt;top:0pt;margin-left:0pt;margin-top:0pt;' +
        `height:${model.style.height};width:${model.style.width};`,
      'o:spid': `_x0000_s${1025 + index}`,
      alt: model.imagedata.title,
      type: '#_x0000_t75',
      'o:spt': '75',
      filled: 'f',
      'o:preferrelative': 't',
      stroked: 'f',
      coordsize: '21600,21600',
    };
  }

  return {
    // id
    id: `_x0000_s${1025 + index}`,
    'o:spid': `_x0000_s${1025 + index}`,
    type: '#_x0000_t202',
    style:
      'position:absolute; margin-left:105.3pt;margin-top:10.5pt;width:97.8pt;height:59.1pt;z-index:1;visibility:hidden',
    fillcolor: 'infoBackground [80]',
    strokecolor: 'none [81]',
    'o:insetmode': model.note && model.note.margins && model.note.margins.insetmode,
  };
};

module.exports = VShapeXform;
