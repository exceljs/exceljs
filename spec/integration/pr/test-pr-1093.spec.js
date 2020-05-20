const proxyquire = require('proxyquire');
const fs = require('fs');
const path = require('path');

const ExcelJS = proxyquire('../../../lib/exceljs.nodejs.js', {
  tmp: {
    file: cb => {
      setTimeout(() => {
        cb(new Error('the error'));
      }, 50);
    },
    '@global': true,
  },
});

// todo Disabled until we drop node 8, since proxyquire breaks readable-streams
describe.skip('pr 1093', () => {
  it('Should fail with error', done => {
    const stream = fs.createReadStream(
      path.join(__dirname, '../data/test-pr-1093.xlsx')
    );

    const wb = new ExcelJS.stream.xlsx.WorkbookReader();
    wb.read(stream, {
      entries: 'emit',
      sharedStrings: 'cache',
      worksheets: 'emit',
    });

    wb.on('end', () => done(new Error('unreachable')));
    wb.on('error', e => e.message === 'the error' && done());
  });
});
