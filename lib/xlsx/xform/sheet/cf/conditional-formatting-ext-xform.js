const _ = require('../../../../utils/under-dash');
const utils = require('../../../../utils/utils');
const colCache = require('../../../../utils/col-cache');
const BaseXform = require('../../base-xform');
const Range = require('../../../../doc/range');

// formatting rules appear to be directly in <worksheet>

// custom - by formula
//   <conditionalFormatting sqref="A1:E5">
//     <cfRule type="expression" dxfId="0" priority="1">
//       <formula>IF(MOD(ROW()+COLUMN(),2)=0,TRUE,FALSE)</formula>
//     </cfRule>
//   </conditionalFormatting>
//
// greaterThan
//   <conditionalFormatting sqref="A1:E6">
//     <cfRule type="cellIs" dxfId="6" priority="1" operator="greaterThan">
//       <formula>13</formula>
//     </cfRule>
//   </conditionalFormatting>
// top 10%
//   <conditionalFormatting sqref="A1:E6">
//     <cfRule type="top10" dxfId="5" priority="1" percent="1" rank="10"/>
//   </conditionalFormatting>
// bottom 10
//   <conditionalFormatting sqref="A1:E6">
//     <cfRule type="top10" dxfId="4" priority="1" bottom="1" rank="10"/>
//   </conditionalFormatting>
// gradient fill
//   Needs some study
// solid fill
//   like gradient fill
// Green Yellow Red (looks a bit like gradient path)
//   <conditionalFormatting sqref="A1:E6">
//     <cfRule type="colorScale" priority="1">
//       <colorScale>
//         <cfvo type="min"/>
//         <cfvo type="percentile" val="50"/>
//         <cfvo type="max"/>
//         <color rgb="FFF8696B"/>
//         <color rgb="FFFFEB84"/>
//         <color rgb="FF63BE7B"/>
//       </colorScale>
//     </cfRule>
//   </conditionalFormatting>
// Arrow Icons
//   <conditionalFormatting sqref="A1:E6">
//     <cfRule type="iconSet" priority="1">
//       <iconSet iconSet="3Arrows">
//         <cfvo type="percent" val="0"/>
//         <cfvo type="percent" val="33"/>
//         <cfvo type="percent" val="67"/>
//       </iconSet>
//     </cfRule>
//   </conditionalFormatting>
// Star Icons - needs more work
// Equal to 13
//   <conditionalFormatting sqref="A1:E6">
//     <cfRule type="cellIs" dxfId="3" priority="1" operator="equal">
//       <formula>13</formula>
//     </cfRule>
//   </conditionalFormatting>
// Text Contains
//   <conditionalFormatting sqref="A1:E6">
//     <cfRule type="containsText" dxfId="0" priority="1" operator="containsText" text="sday">
//       <formula>NOT(ISERROR(SEARCH("sday",A1)))</formula>
//     </cfRule>
//   </conditionalFormatting>
// Multiple (fully overlapping)
//   <conditionalFormatting sqref="A1:E6">
//     <cfRule type="top10" dxfId="2" priority="2" percent="1" rank="10"/>
//     <cfRule type="top10" dxfId="1" priority="1" percent="1" bottom="1" rank="10"/>
//   </conditionalFormatting>
// Intersecting
//   <conditionalFormatting sqref="A1:C6">
//     <cfRule type="top10" dxfId="2" priority="2" percent="1" bottom="1" rank="10"/>
//   </conditionalFormatting>
//   <conditionalFormatting sqref="C1:E6">
//     <cfRule type="top10" dxfId="0" priority="1" percent="1" rank="10"/>
//   </conditionalFormatting>
// Infinite range
//   <conditionalFormatting sqref="A3:XFD5">
//     <cfRule type="cellIs" dxfId="29" priority="2" operator="between">
//       <formula>11</formula>
//       <formula>15</formula>
//     </cfRule>
//   </conditionalFormatting>
//   <conditionalFormatting sqref="B1:D1048576">
//     <cfRule type="expression" dxfId="28" priority="1">
//       <formula>IF(MOD(B1,5)=3,TRUE,FALSE)</formula>
//     </cfRule>
//   </conditionalFormatting>
// Special Dates
//   <conditionalFormatting sqref="A1:G6">
//     <cfRule type="timePeriod" dxfId="0" priority="12" timePeriod="thisWeek">
//       <formula>AND(TODAY()-ROUNDDOWN(A1,0)&lt;=WEEKDAY(TODAY())-1,ROUNDDOWN(A1,0)-TODAY()&lt;=7-WEEKDAY(TODAY()))
//       </formula>
//     </cfRule>
//     <cfRule type="timePeriod" dxfId="1" priority="11" timePeriod="lastWeek">
//       <formula>AND(TODAY()-ROUNDDOWN(A1,0)&gt;=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(A1,0)&lt;(WEEKDAY(TODAY())+7))
//       </formula>
//     </cfRule>
//     <cfRule type="timePeriod" dxfId="2" priority="10" timePeriod="nextWeek">
//       <formula>AND(ROUNDDOWN(A1,0)-TODAY()&gt;(7-WEEKDAY(TODAY())),ROUNDDOWN(A1,0)-TODAY()&lt;(15-WEEKDAY(TODAY())))
//       </formula>
//     </cfRule>
//     <cfRule type="timePeriod" dxfId="3" priority="9" timePeriod="yesterday">
//       <formula>FLOOR(A1,1)=TODAY()-1</formula>
//     </cfRule>
//     <cfRule type="timePeriod" dxfId="4" priority="8" timePeriod="today">
//       <formula>FLOOR(A1,1)=TODAY()</formula>
//     </cfRule>
//     <cfRule type="timePeriod" dxfId="5" priority="6" timePeriod="tomorrow">
//       <formula>FLOOR(A1,1)=TODAY()+1</formula>
//     </cfRule>
//     <cfRule type="timePeriod" dxfId="6" priority="5" timePeriod="last7Days">
//       <formula>AND(TODAY()-FLOOR(A1,1)&lt;=6,FLOOR(A1,1)&lt;=TODAY())</formula>
//     </cfRule>
//     <cfRule type="timePeriod" dxfId="7" priority="3" timePeriod="lastMonth">
//       <formula>AND(MONTH(A1)=MONTH(EDATE(TODAY(),0-1)),YEAR(A1)=YEAR(EDATE(TODAY(),0-1)))</formula>
//     </cfRule>
//     <cfRule type="timePeriod" dxfId="8" priority="2" timePeriod="thisMonth">
//       <formula>AND(MONTH(A1)=MONTH(TODAY()),YEAR(A1)=YEAR(TODAY()))</formula>
//     </cfRule>
//     <cfRule type="timePeriod" dxfId="9" priority="1" timePeriod="nextMonth">
//       <formula>AND(MONTH(A1)=MONTH(EDATE(TODAY(),0+1)),YEAR(A1)=YEAR(EDATE(TODAY(),0+1)))</formula>
//     </cfRule>
//   </conditionalFormatting>


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

function optimiseDataValidations(model) {
  // Squeeze alike data validations together into rectangular ranges
  // to reduce file size and speed up Excel load time
  const dvList = _
    .map(model, (dataValidation, address) => ({
      address,
      dataValidation,
      marked: false,
    }))
    .sort((a, b) => _.strcmp(a.address, b.address));
  const dvMap = _.keyBy(dvList, 'address');
  const matchCol = (addr, height, col) => {
    for (let i = 0; i < height; i++) {
      const otherAddress = colCache.encodeAddress(addr.row + i, col);
      if (!model[otherAddress] || !_.isEqual(model[addr.address], model[otherAddress])) {
        return false;
      }
    }
    return true;
  };
  return dvList
    .map(dv => {
      if (!dv.marked) {
        const addr = colCache.decodeAddress(dv.address);

        // iterate downwards - finding matching cells
        let height = 1;
        let otherAddress = colCache.encodeAddress(addr.row + height, addr.col);
        while (model[otherAddress] && _.isEqual(dv.dataValidation, model[otherAddress])) {
          height++;
          otherAddress = colCache.encodeAddress(addr.row + height, addr.col);
        }

        // iterate rightwards...

        let width = 1;
        while (matchCol(addr, height, addr.col + width)) {
          width++;
        }

        // mark all included addresses
        for (let i = 0; i < height; i++) {
          for (let j = 0; j < width; j++) {
            otherAddress = colCache.encodeAddress(addr.row + i, addr.col + j);
            dvMap[otherAddress].marked = true;
          }
        }

        if ((height > 1) || (width > 1)) {
          const bottom = addr.row + (height - 1);
          const right = addr.col + (width - 1);
          return {
            ...dv.dataValidation,
            sqref: `${dv.address}:${colCache.encodeAddress(bottom, right)}`,
          };
        }
        return {
          ...dv.dataValidation,
          sqref: dv.address,
        };
      }
      return null;
    })
    .filter(Boolean);
}

class DataValidationsXform extends BaseXform {
  get tag() {
    return 'dataValidations';
  }

  render(xmlStream, model) {
    const optimizedModel = optimiseDataValidations(model);
    if (optimizedModel.length) {
      xmlStream.openNode('dataValidations', {count: optimizedModel.length});

      optimizedModel.forEach(value => {
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
        xmlStream.addAttribute('sqref', value.sqref);
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
        const dataValidation = node.attributes.type ?
          {type: node.attributes.type, formulae: []} :
          {type: 'any'};

        if (node.attributes.type) {
          assignBool(dataValidation, node.attributes, 'allowBlank');
        }
        assignBool(dataValidation, node.attributes, 'showInputMessage');
        assignBool(dataValidation, node.attributes, 'showErrorMessage');

        switch (dataValidation.type) {
          case 'any':
          case 'list':
          case 'custom':
            break;
          default:
            assign(dataValidation, node.attributes, 'operator', 'between');
            break;
        }
        assign(dataValidation, node.attributes, 'promptTitle');
        assign(dataValidation, node.attributes, 'prompt');
        assign(dataValidation, node.attributes, 'errorStyle');
        assign(dataValidation, node.attributes, 'errorTitle');
        assign(dataValidation, node.attributes, 'error');

        this._dataValidation = dataValidation;
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
    if (this._formula) {
      this._formula.push(text);
    }
  }

  parseClose(name) {
    switch (name) {
      case 'dataValidations':

        return false;
      case 'dataValidation':
        if (!this._dataValidation.formulae || !this._dataValidation.formulae.length) {
          delete this._dataValidation.formulae;
          delete this._dataValidation.operator;
        }
        if (this._address.includes(':')) {
          const range = new Range(this._address);
          range.forEachAddress(address => {
            this.model[address] = this._dataValidation;
          });
        } else {
          this.model[this._address] = this._dataValidation;
        }
        return true;
      case 'formula1':
      case 'formula2': {
        let formula = this._formula.join('');
        switch (this._dataValidation.type) {
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
        this._dataValidation.formulae.push(formula);
        this._formula = undefined;
        return true;
      }
      default:
        return true;
    }
  }
}

module.exports = DataValidationsXform;
