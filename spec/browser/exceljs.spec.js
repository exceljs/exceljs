/* global ExcelJS */
// ExcelJS is a global injected by `./dist/exceljs.js` during jasmine's setup

'use strict';

function unexpectedError(done) {
  return function(error) {
    // eslint-disable-next-line no-console
    console.error('Error Caught', error.message, error.stack);
    expect(true).toEqual(false);
    done();
  };
}

describe('ExcelJS', () => {
  it('should read and write xlsx via binary buffer', done => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('A2').value = 7;

    wb.xlsx
      .writeBuffer()
      .then(buffer => {
        const wb2 = new ExcelJS.Workbook();
        return wb2.xlsx.load(buffer).then(() => {
          const ws2 = wb2.getWorksheet('blort');
          expect(ws2).toBeTruthy();

          expect(ws2.getCell('A1').value).toEqual('Hello, World!');
          expect(ws2.getCell('A2').value).toEqual(7);
          done();
        });
      })
      .catch(error => {
        throw error;
      })
      .catch(unexpectedError(done));
  });
  it('should read xlsx via blob', async done => {

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('A2').value = 7;

    wb.xlsx
            .writeBuffer()
            .then(async buffer => {
              // eslint-disable-next-line no-undef
              const blob = new Blob([buffer], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
              });
              const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(blob);
              let rows = 0;
              for await (const worksheetReader of workbookReader) {
                expect(worksheetReader.name).toEqual('blort');
                for await (const row of worksheetReader) {
                  rows++;
                  if (rows === 1) {
                    expect(row.getCell('A').value).toEqual('Hello, World!');
                  } else if (rows === 2) {
                    expect(row.getCell('A').value).toEqual(7);
                  }
                }
              }
              done();
            })
            .catch(error => {
              throw error;
            })
            .catch(unexpectedError(done));
  });
  it('should read and write xlsx via base64 buffer', done => {
    const options = {
      base64: true,
    };
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('A2').value = 7;

    wb.xlsx
      .writeBuffer(options)
      .then(buffer => {
        const wb2 = new ExcelJS.Workbook();
        return wb2.xlsx.load(buffer.toString('base64'), options).then(() => {
          const ws2 = wb2.getWorksheet('blort');
          expect(ws2).toBeTruthy();

          expect(ws2.getCell('A1').value).toEqual('Hello, World!');
          expect(ws2.getCell('A2').value).toEqual(7);
          done();
        });
      })
      .catch(error => {
        throw error;
      })
      .catch(unexpectedError(done));
  });
  it('should write csv via buffer', done => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');

    ws.getCell('A1').value = 'Hello, World!';
    ws.getCell('B1').value = 'What time is it?';
    ws.getCell('A2').value = 7;
    ws.getCell('B2').value = '12pm';

    wb.csv
      .writeBuffer()
      .then(buffer => {
        expect(buffer.toString()).toEqual(
          '"Hello, World!",What time is it?\n7,12pm'
        );
        done();
      })
      .catch(error => {
        throw error;
      })
      .catch(unexpectedError(done));
  });
});
