
const BaseXform = require('exceljs/lib/xlsx/xform/base-xform');
const OffXform = require('./off');
const ExtXform = require('./ext');
class XFrm extends BaseXform {
   constructor() {
     super();

     this.map = {
       'a:ext': new ExtXform(),
       'a:off': new OffXform(),
     };
   }
  get tag() {
    return 'a:xfrm';
  }

  render(xmlStream, model) {
    
    xmlStream.openNode(this.tag, model.rotation && {
      'rot': (model.rotation * 60000),
    });
    this.map['a:ext'].render(xmlStream, model);
    this.map['a:off'].render(xmlStream, model);
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
          rot: parseInt(node.attributes['rot']) / 60000,
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
           this.model.ext = this.map['a:ext'].model;
           this.model.off = this.map['a:off'].model;
           return false;
         default:
           return true;
       }
       }
}

module.exports = XFrm;