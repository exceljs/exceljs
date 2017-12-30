'use strict';

var setConfigValue = require('./config/set-value');
setConfigValue('promise', require('promish'), false);

var ExcelJS = {
  Workbook: require('./doc/workbook'),
  ModelContainer: require('./doc/modelcontainer'),
  stream: {
    xlsx: {
      WorkbookWriter: require('./stream/xlsx/workbook-writer'),
      WorkbookReader: require('./stream/xlsx/workbook-reader')
    }
  },
  config: {
    setValue: require('./config/set-value')
  }
};

Object.assign(ExcelJS, require('./doc/enums'));

module.exports = ExcelJS;
//# sourceMappingURL=exceljs.nodejs.js.map
