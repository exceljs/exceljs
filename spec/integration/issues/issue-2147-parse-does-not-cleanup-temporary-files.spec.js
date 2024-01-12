const {expect} = require('chai');
const fs = require('fs');
const tmp = require('tmp');

const ExcelJS = verquire('exceljs');

// test file - any excel file works for this test
const TEST_XLSX_FILE_NAME = './spec/integration/data/test-row-styles.xlsx';

describe('github issues', () => {
  let origTmpFile = null;
  let cleanupCalled = false;

  before(() => {
    // hooking tmp.file to check if the cleanupCallback
    // gets invoked during the test
    origTmpFile = tmp.file;

    tmp.file = function(cb) {
      origTmpFile(function(err, path, fd, cleanupCallback) {
        cb(err, path, fd, function() {
          cleanupCalled = true;
          cleanupCallback();
        });
      });
    };
  });

  after(() => {
    tmp.file = origTmpFile;
    origTmpFile = null;
    cleanupCalled = false;
  });

  it('issue 2147 - parse() doesn\'t cleanup temporary files if not fully iterated', async () => {
    const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(
      fs.createReadStream(TEST_XLSX_FILE_NAME)
    );
    // eslint-disable-next-line no-unused-vars
    for await (const _ of workbookReader) {
      // aborting iteration prematurely
      // throw or return would have the same effect
      break;
    }

    expect(cleanupCalled).to.be.true(
      'tmp.file cleanupCallback has not been called'
    );
  });
});
