/* eslint-disable import/no-extraneous-dependencies,node/no-unpublished-require */
import 'core-js/modules/es.promise'
import 'core-js/modules/es.promise.finally'
import 'core-js/modules/es.object.assign'
import 'core-js/modules/es.object.keys'
import 'core-js/modules/es.object.values'
import 'core-js/modules/es.symbol'
import 'core-js/modules/es.symbol.async-iterator'
// required by core-js/modules/es.promise Promise.all
import 'core-js/modules/es.array.iterator'
// required by node_modules/saxes/saxes.js SaxesParser.captureTo
import 'core-js/modules/es.array.includes'
// required by lib/doc/workbook.js Workbook.model
import 'core-js/modules/es.array.find-index'
// required by lib/doc/workbook.js Workbook.addWorksheet and Workbook.getWorksheet
import 'core-js/modules/es.array.find'
// required by node_modules/saxes/saxes.js SaxesParser.getCode10
import 'core-js/modules/es.string.from-code-point'
// required by lib/xlsx/xform/sheet/data-validations-xform.js DataValidationsXform.parseClose
import 'core-js/modules/es.string.includes'
// required by lib/utils/utils.js utils.validInt and lib/csv/csv.js CSV.read
import 'core-js/modules/es.number.is-nan'
import 'regenerator-runtime/runtime'
// Object.assign mono-fill
import Enums from './doc/enums'
import wb from './doc/workbook'

const ExcelJS = {
  Workbook: wb
};


Object.keys(Enums).forEach(key => {
  ExcelJS[key] = Enums[key];
});

export default ExcelJS;
