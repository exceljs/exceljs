const ExcelJS = {
  Workbook: require('./lib/doc/workbook'),
  ModelContainer: require('./lib/doc/modelcontainer'),
  stream: {
    xlsx: {
      WorkbookWriter: require('./lib/stream/xlsx/workbook-writer'),
      WorkbookReader: require('./lib/stream/xlsx/workbook-reader'),
    },
  },
};

Object.assign(ExcelJS, require('./lib/doc/enums'));

module.exports = ExcelJS;
