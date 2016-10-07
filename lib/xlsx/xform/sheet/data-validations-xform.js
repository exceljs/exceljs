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
var _ = require('lodash');
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
        xmlStream.addAttribute('type', value.type);
        if (value.operator && (value.type !== 'list') && (value.operator !== 'between')) {
          xmlStream.addAttribute('operator', value.operator);
        }
        if (value.allowBlank) {
          xmlStream.addAttribute('allowBlank', '1');
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
        value.formulae.forEach(function (formula, index) {
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
        var definedName = this._definedName = {
          type: node.attributes.type,
          formulae: []
        };

        assignBool(definedName, node.attributes, 'allowBlank');
        assignBool(definedName, node.attributes, 'showInputMessage');
        assignBool(definedName, node.attributes, 'showErrorMessage');

        switch (definedName.type) {
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


// <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B1">
//   <formula0>Nephews</formula0>
// </dataValidation>


// <dataValidations count="34">
//   <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E1">
//   <formula1>Ducks</formula1>
//   </dataValidation>
//   <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C1">
//   <formula1>Ducks</formula1>
//   </dataValidation>
//   <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C5 E5">
//   <formula1>"Tom, Dick, Harry"</formula1>
//   </dataValidation>
//   <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C7">
//   <formula1>$A$7:$A$9</formula1>
//   </dataValidation>
//   <dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C11">
//   <formula1>5</formula1>
//   <formula2>10</formula2>
//   </dataValidation>
//   <dataValidation type="whole" operator="notBetween" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E11">
//   <formula1>3</formula1>
//   <formula2>7</formula2>
//   </dataValidation>
//   <dataValidation type="whole" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="G11">
//   <formula1>7</formula1>
//   </dataValidation>
//   <dataValidation type="whole" operator="notEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="I11">
//   <formula1>5</formula1>
//   </dataValidation>
//   <dataValidation type="whole" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="K11">
//   <formula1>3</formula1>
//   </dataValidation>
//   <dataValidation type="whole" operator="lessThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="M11">
//   <formula1>6</formula1>
//   </dataValidation>
//   <dataValidation type="whole" operator="greaterThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="O11">
//   <formula1>10</formula1>
//   </dataValidation>
//   <dataValidation type="whole" operator="lessThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="Q11">
//   <formula1>11</formula1>
//   </dataValidation>
//   <dataValidation type="decimal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C13">
//   <formula1>5.5</formula1>
//   <formula2>6.6</formula2>
//   </dataValidation>
//   <dataValidation type="decimal" operator="notBetween" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E13">
//   <formula1>2.2</formula1>
//   <formula2>5.5</formula2>
//   </dataValidation>
//   <dataValidation type="decimal" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="G13">
//   <formula1>3.1415926</formula1>
//   </dataValidation>
//   <dataValidation type="decimal" operator="notEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="I13">
//   <formula1>5</formula1>
//   </dataValidation>
//   <dataValidation type="decimal" operator="greaterThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="K13">
//   <formula1>4</formula1>
//   </dataValidation>
//   <dataValidation type="decimal" operator="lessThan" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="M13">
//   <formula1>77</formula1>
//   </dataValidation>
//   <dataValidation type="decimal" operator="greaterThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="O13">
//   <formula1>4</formula1>
//   </dataValidation>
//   <dataValidation type="decimal" operator="lessThanOrEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="Q13">
//   <formula1>4</formula1>
//   </dataValidation>
//   <dataValidation type="whole" showInputMessage="1" showErrorMessage="1" sqref="S11">
//   <formula1>5</formula1>
//   <formula2>7</formula2>
//   </dataValidation>
//   <dataValidation type="decimal" showInputMessage="1" showErrorMessage="1" sqref="S13">
//   <formula1>1</formula1>
//   <formula2>6</formula2>
//   </dataValidation>
//   <dataValidation type="date" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C15">
//   <formula1>42370</formula1>
//   <formula2>42400</formula2>
//   </dataValidation>
//   <dataValidation type="textLength" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C19">
//   <formula1>5</formula1>
//   <formula2>10</formula2>
//   </dataValidation>
//   <dataValidation type="textLength" operator="notBetween" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E19">
//   <formula1>5</formula1>
//   <formula2>10</formula2>
//   </dataValidation>
//   <dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C21">
//   <formula1>OR(C21=5,C21=7)</formula1>
//   </dataValidation>
//   <dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="This is Title" prompt="This is Input Message" sqref="C23">
//   <formula1>3</formula1>
//   <formula2>8</formula2>
//   </dataValidation>
//   <dataValidation type="whole" operator="equal" allowBlank="1" showErrorMessage="1" sqref="C25">
//   <formula1>5</formula1>
//   </dataValidation>
//   <dataValidation type="whole" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Error Title" error="Error Message" sqref="C27">
//   <formula1>5</formula1>
//   </dataValidation>
//   <dataValidation type="whole" errorStyle="warning" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Warning Title" error="Warning Message" sqref="E27">
//   <formula1>5</formula1>
//   </dataValidation>
//   <dataValidation type="whole" errorStyle="information" operator="equal" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Info Title" error="Info Message" sqref="G27">
//   <formula1>5</formula1>
//   </dataValidation>
//   <dataValidation type="whole" operator="equal" allowBlank="1" showInputMessage="1" errorTitle="Don't Show Title" error="Don't Show Message" sqref="I27">
//   <formula1>5</formula1>
//   </dataValidation>
//   <dataValidation type="whole" operator="equal" allowBlank="1" showErrorMessage="1" promptTitle="Don't Show Title" prompt="Don't Show Message" sqref="E23">
//   <formula1>5</formula1>
//   </dataValidation>
//   <dataValidation type="whole" operator="notEqual" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C31">
//   <formula1>odds</formula1>
//   </dataValidation>
// </dataValidations>
