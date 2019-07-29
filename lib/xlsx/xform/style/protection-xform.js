const BaseXform = require('../base-xform');

const validation = {
  locked(value) {
    return value ? true : undefined;
  },
};

// Protection encapsulates translation from style.protection model to/from xlsx
class ProtectionXform extends BaseXform {
  get tag() {
    return 'protection';
  }

  render(xmlStream, model) {
    xmlStream.addRollback();
    xmlStream.openNode('protection');

    let isValid = false;
    function add(name, value) {
      if (value) {
        xmlStream.addAttribute(name, value);
        isValid = true;
      }
    }
    add('locked', validation.locked(!model.locked) ? '0' : false);

    xmlStream.closeNode();

    if (isValid) {
      xmlStream.commit();
    } else {
      xmlStream.rollback();
    }
  }

  parseOpen(node) {
    const model = {};

    let valid = false;
    function add(truthy, name, value) {
      if (truthy) {
        model[name] = value;
        valid = true;
      }
    }
    add(node.attributes.locked, 'locked', !!node.attributes.locked);
    this.model = valid ? model : null;
  }

  parseText() {}

  parseClose() {
    return false;
  }
}

module.exports = ProtectionXform;
