const CompositeXform = require('../../composite-xform');

const FormulaExtXform = require('./formula-ext-xform');
const SqRefExtXform = require('../cf-ext/sqref-ext-xform');

function assign(definedName, attributes, name, defaultValue) {
  const value = attributes[name];
  if (value !== undefined) {
    definedName[name] = value;
  } else if (defaultValue !== undefined) {
    definedName[name] = defaultValue;
  }
}
function parseBool(value) {
  switch (value) {
    case '1':
    case 'true':
      return true;
    default:
      return false;
  }
}
function assignBool(definedName, attributes, name, defaultValue) {
  const value = attributes[name];
  if (value !== undefined) {
    definedName[name] = parseBool(value);
  } else if (defaultValue !== undefined) {
    definedName[name] = defaultValue;
  }
}

class DataValidationExtXform extends CompositeXform {
  // Inspired from the data-validation-xform file
  // Normally there should only be lists types,
  // with a formula and a sqRef (as a node, contrary to normal data validations)

  constructor() {
    super();

    this.map = {
      'xm:sqref': (this.sqRef = new SqRefExtXform()),
      'x14:formula1': (this.formula = new FormulaExtXform()),
    };
  }

  get tag() {
    return 'x14:dataValidation';
  }

  static isExt(dataValidation) {
    if (
      dataValidation.formulae &&
      dataValidation.formulae.length === 1 &&
      typeof dataValidation.formulae[0] === 'string'
    ) {
      const worksheetNameRegex = new RegExp(
        [
          /(('[^/\\?*[\]]{1,31}'|[A-Za-z0-9_]{1,31})!)/,
          /((\$?[A-Za-z]{1,3})(\$?[0-9]{1,6}))(:((\$?[A-Za-z]{1,3})(\$?[0-9]{1,6})))?/,
        ]
          .map(reg => reg.source)
          .join('')
      );
      const match = worksheetNameRegex.test(dataValidation.formulae[0]);
      return dataValidation.type === 'list' && match;
    }
    return false;
  }

  render(xmlStream, model) {
    xmlStream.openNode(this.tag);

    // We keep the same logic as with normal data validations,
    // although there should be only the type "list" for external data validations
    if (model.type !== 'any') {
      xmlStream.addAttribute('type', model.type);

      if (model.allowBlank) {
        xmlStream.addAttribute('allowBlank', '1');
      }
    }
    if (model.showInputMessage) {
      xmlStream.addAttribute('showInputMessage', '1');
    }
    if (model.promptTitle) {
      xmlStream.addAttribute('promptTitle', model.promptTitle);
    }
    if (model.prompt) {
      xmlStream.addAttribute('prompt', model.prompt);
    }
    if (model.showErrorMessage) {
      xmlStream.addAttribute('showErrorMessage', '1');
    }
    if (model.errorStyle) {
      xmlStream.addAttribute('errorStyle', model.errorStyle);
    }
    if (model.errorTitle) {
      xmlStream.addAttribute('errorTitle', model.errorTitle);
    }
    if (model.error) {
      xmlStream.addAttribute('error', model.error);
    }

    // First the formula then the sqref
    this.formula.render(xmlStream, model.formulae);
    this.sqRef.render(xmlStream, model.sqref);

    xmlStream.closeNode();
  }

  createNewModel(node) {
    const dataValidation = {type: node.attributes.type || 'list'};

    if (node.attributes.type) {
      assignBool(dataValidation, node.attributes, 'allowBlank');
    }
    assignBool(dataValidation, node.attributes, 'showInputMessage');
    assignBool(dataValidation, node.attributes, 'showErrorMessage');
    assign(dataValidation, node.attributes, 'promptTitle');
    assign(dataValidation, node.attributes, 'prompt');
    assign(dataValidation, node.attributes, 'errorStyle');
    assign(dataValidation, node.attributes, 'errorTitle');
    assign(dataValidation, node.attributes, 'error');

    return dataValidation;
  }

  onParserClose(name, parser) {
    switch (name) {
      case 'xm:sqref':
        this.model._address = parser.model;
        break;

      case 'x14:formula1':
        this.model.formulae = [parser.model['xm:f']];
        break;
    }
  }
}

module.exports = DataValidationExtXform;
