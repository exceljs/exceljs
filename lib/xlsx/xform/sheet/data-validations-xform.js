const _ = require('../../../utils/under-dash');
const utils = require('../../../utils/utils');
const BaseXform = require('../base-xform');

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

function mergeDataValidations(model) {
  const valueAddressesMap = new Map();
  _.each(model, (value, address) => {
    const key = JSON.stringify(value);
    let mapPerColumn = valueAddressesMap.get(key);
    if (!mapPerColumn) {
      mapPerColumn = new Map();
      valueAddressesMap.set(key, mapPerColumn);
    }

    const columnNo = address.match(/[A-Z]+/)[0];
    const rowNo = parseInt(address.match(/\d+/)[0], 10);

    let rowArray = mapPerColumn.get(columnNo);
    if (!rowArray) {
      rowArray = [];
      mapPerColumn.set(columnNo, rowArray);
    }
    rowArray.push(rowNo);
  });

  const mergedResults = {};
  valueAddressesMap.forEach((columnsMap, formulaeValue) => {
    columnsMap.forEach((rows, columnNo) => {
      const sortedRowNumbers = rows.sort((a, b) => a - b);
      const arrayOfRanges = [];
      sortedRowNumbers.forEach(rowNum => {
        const previousNumber = arrayOfRanges.length > 0 ? arrayOfRanges[arrayOfRanges.length - 1].max : null;
        if (previousNumber && previousNumber === rowNum - 1) {
          arrayOfRanges[arrayOfRanges.length - 1].max = rowNum;
        } else {
          arrayOfRanges.push({min: rowNum, max: rowNum});
        }
      });
      arrayOfRanges.forEach(range => {
        if (range.min === range.max) {
          mergedResults[`${columnNo}${range.min}`] = JSON.parse(formulaeValue);
        } else {
          mergedResults[`${columnNo}${range.min}:${columnNo}${range.max}`] = JSON.parse(formulaeValue);
        }
      });
    });
  });

  return mergedResults;
}

class DataValidationsXform extends BaseXform {
  get tag() {
    return 'dataValidations';
  }

  render(xmlStream, model) {
    const optimizedModel = mergeDataValidations(model);
    const count = optimizedModel && Object.keys(optimizedModel).length;
    if (count) {
      xmlStream.openNode('dataValidations', {count});

      _.each(optimizedModel, (value, address) => {
        xmlStream.openNode('dataValidation');
        if (value.type !== 'any') {
          xmlStream.addAttribute('type', value.type);

          if (value.operator && value.type !== 'list' && value.operator !== 'between') {
            xmlStream.addAttribute('operator', value.operator);
          }
          if (value.allowBlank) {
            xmlStream.addAttribute('allowBlank', '1');
          }
        }
        if (value.showInputMessage) {
          xmlStream.addAttribute('showInputMessage', '1');
        }
        if (value.promptTitle) {
          xmlStream.addAttribute('promptTitle', value.promptTitle);
        }
        if (value.prompt) {
          xmlStream.addAttribute('prompt', value.prompt);
        }
        if (value.showErrorMessage) {
          xmlStream.addAttribute('showErrorMessage', '1');
        }
        if (value.errorStyle) {
          xmlStream.addAttribute('errorStyle', value.errorStyle);
        }
        if (value.errorTitle) {
          xmlStream.addAttribute('errorTitle', value.errorTitle);
        }
        if (value.error) {
          xmlStream.addAttribute('error', value.error);
        }
        xmlStream.addAttribute('sqref', address);
        (value.formulae || []).forEach((formula, index) => {
          xmlStream.openNode(`formula${index + 1}`);
          if (value.type === 'date') {
            xmlStream.writeText(utils.dateToExcel(new Date(formula)));
          } else {
            xmlStream.writeText(formula);
          }
          xmlStream.closeNode();
        });
        xmlStream.closeNode();
      });
      xmlStream.closeNode();
    }
  }

  parseOpen(node) {
    switch (node.name) {
      case 'dataValidations':
        this.model = {};
        return true;

      case 'dataValidation': {
        this._address = node.attributes.sqref;
        const definedName = node.attributes.type
          ? {
            type: node.attributes.type,
            formulae: [],
          }
          : {
            type: 'any',
          };

        if (node.attributes.type) {
          assignBool(definedName, node.attributes, 'allowBlank');
        }
        assignBool(definedName, node.attributes, 'showInputMessage');
        assignBool(definedName, node.attributes, 'showErrorMessage');

        switch (definedName.type) {
          case 'any':
          case 'list':
          case 'custom':
            break;
          default:
            assign(definedName, node.attributes, 'operator', 'between');
            break;
        }
        assign(definedName, node.attributes, 'promptTitle');
        assign(definedName, node.attributes, 'prompt');
        assign(definedName, node.attributes, 'errorStyle');
        assign(definedName, node.attributes, 'errorTitle');
        assign(definedName, node.attributes, 'error');

        this._definedName = definedName;
        return true;
      }

      case 'formula1':
      case 'formula2':
        this._formula = [];
        return true;

      default:
        return false;
    }
  }

  parseText(text) {
    this._formula.push(text);
  }

  parseClose(name) {
    switch (name) {
      case 'dataValidations':
        return false;
      case 'dataValidation':
        if (!this._definedName.formulae || !this._definedName.formulae.length) {
          delete this._definedName.formulae;
          delete this._definedName.operator;
        }
        this.model[this._address] = this._definedName;
        return true;
      case 'formula1':
      case 'formula2': {
        let formula = this._formula.join('');
        switch (this._definedName.type) {
          case 'whole':
          case 'textLength':
            formula = parseInt(formula, 10);
            break;
          case 'decimal':
            formula = parseFloat(formula);
            break;
          case 'date':
            formula = utils.excelToDate(parseFloat(formula));
            break;
          default:
            break;
        }
        this._definedName.formulae.push(formula);
        return true;
      }
      default:
        return true;
    }
  }
}

module.exports = DataValidationsXform;
