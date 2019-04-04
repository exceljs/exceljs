const fs = require('fs');
const _ = require('../lib/utils/under-dash.js');
const Relationships = require('../lib/doc/relationships');

const testUtils = require('./utils/utils');

const mode = process.argv[2];
const out = process.argv[3];
const filename = process.argv[4];

let rt;
const tests = {
  document() {
    rt = new Relationships({
      filename: '',
      path: '',
      templatename: '',
      templatepath: '',
    });
    rt.addOfficeDocument();
  },
  sheet() {
    rt = new Relationships({
      filename: 'sheet1.xml',
      path: `${out}/xl/worksheets`,
      templatename: 'sheet.xml',
      templatepath: '../templates/xl/worksheets',
    });

    const lst = ['http://www.google.com', 'https://www.facebook.com'];
    _.each(lst, item => {
      rt.addHyperlink(item);
    });
  },
  book() {
    rt = new Relationships({
      filename: 'workbook.xml',
      path: `${out}/xl`,
      templatename: 'workbook.xml',
      templatepath: '../templates/xl',
    });

    rt.addWorksheet(1);
    rt.addTheme(1);
    rt.addStyles();
    rt.addSharedStrings();
  },
};

testUtils
  .cleanDir(out)
  .then(() => {
    // init test
    tests[mode]();

    // serialize
    const fullpath = `${out}/${filename}`;
    console.log(`Writing relationship table to ${fullpath}`);
    const stream = fs.createWriteStream(fullpath);
    rt.write(stream).then(() => {
      stream.close();
      console.log('Done.');
    });
  })
  .catch(err => {
    console.log(`Caught Error: ${err}`);
  });
