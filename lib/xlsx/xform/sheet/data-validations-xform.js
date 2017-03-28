/**
 * Copyright (c) 2016 Guyon Roche
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 */
'use strict';
var _ = require('../../../utils/under-dash');
var utils = require('../../../utils/utils');
var BaseXform = require('../base-xform');

// for dev reference
// var DataValidationTemplate = { // a _ denotes defaults
//   address: 'A1',
//   type: ['list', 'whole', 'decimal', 'date', 'textLength'],
//   operator: ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'],
//   allowBlank: [true, false],
//   showInputMessage: [true, false],
//   showErrorMessage: [true, false],
//   formulae: [['DefinedName', '<value>', '"list,of,values"', '$A$1', '$A$1:$B$2','ACTUAL("Formula")']],
//   promptTitle: ['Title'],
//   prompt: ['words'],
//   errorStyle: ['error', 'warning', 'information'],
//   errorTitle: ['Title'],
//   error: ['words']
// };


function assign(definedName, attributes, name, defaultValue) {
  var value = attributes[name];
  if (value !== undefined) {
    definedName[name] = value;
  } else if (defaultValue !== undefined) {
    definedName[name] = defaultValue;
  }
}
function parseBool(value) {
  switch(value) {
    case '1':
    case 'true':
      return true;
    default:
      return false;
  }
}
function assignBool(definedName, attributes, name, defaultValue) {
  var value = attributes[name];
  if (value !== undefined) {
    definedName[name] = parseBool(value);
  } else if (defaultValue !== undefined) {
    definedName[name] = defaultValue;
  }
}

var DataValidationsXform = module.exports = function() {
};

utils.inherits(DataValidationsXform, BaseXform, {

  get tag() { return 'dataValidations'; },

  render: function(xmlStream, model) {
    var count = model && Object.keys(model).length;
    if (count) {
      xmlStream.openNode('dataValidations', {count: count});

      _.each(model, function (value, address) {
        xmlStream.openNode('dataValidation');
        if (value.type !== 'any') {
          xmlStream.addAttribute('type', value.type);

          if (value.operator && (value.type !== 'list') && (value.operator !== 'between')) {
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
        (value.formulae || []).forEach(function (formula, index) {
          xmlStream.openNode('formula' + (index + 1));
          if (value.type === 'date') {
            xmlStream.writeText(utils.dateToExcel(formula));
          } else {
            xmlStream.writeText(formula);
          }
          xmlStream.closeNode();
        });
        xmlStream.closeNode();
      });
      xmlStream.closeNode();
    }
  },
  parseOpen: function(node) {
    switch(node.name) {
      case 'dataValidations':
        this.model = {};
        return true;

      case 'dataValidation':
        this._address = node.attributes.sqref;
        var definedName = this._definedName = node.attributes.type ? {
            type: node.attributes.type,
            formulae: []
          } : {
            type: 'any'
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
        return true;
      case 'formula1':
      case 'formula2':
        this._formula = [];
        return true;
      
      default:
        return false;
    }
  },
  parseText: function(text) {
    this._formula.push(text);
  },
  parseClose: function(name) {
    switch(name) {
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
      case 'formula2':
        var formula = this._formula.join('');
        switch (this._definedName.type) {
          case 'whole':
          case 'textLength':
            formula = parseInt(formula);
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
      default:
        return true;
    }
  }
});
