const ExcelJS = {
  Workbook: require('./doc/workbook'),
  ModelContainer: require('./doc/modelcontainer'),
  stream: {
    xlsx: {
      WorkbookWriter: require('./stream/xlsx/workbook-writer'),
      WorkbookReader: require('./stream/xlsx/workbook-reader'),
    },
  },
  // [exceljs-hardened] Use utils.safeJoin(baseDir, userInput) instead of
  // Node's path.join() wherever you build an addImage({filename}) path
  // from user-influenced input. See lib/utils/media-path-guard.js for why
  // path.join() alone cannot be safely validated after the fact.
  utils: {
    safeJoin: require('./utils/media-path-guard').safeJoin,
  },
};

Object.assign(ExcelJS, require('./doc/enums'));

module.exports = ExcelJS;
