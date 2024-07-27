const {expect} = require('chai');

const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  describe('issue-2789 Hidden', async () => {
    const fileList = [
      'google-sheets',
      'libre-calc-as-excel-2007-365',
      'libre-calc-as-office-open-xml-spreadsheet',
    ];
    for (const file of fileList) {
      it(`Should set hidden attribute correctly (${file})`, async () => {
        const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(
          `./spec/integration/data/hidden-test/${file}.xlsx`
        );
        for await (const worksheetReader of workbookReader) {
          let readRowNum = 1;
          //  Check rows
          for await (const row of worksheetReader) {
            expect(row.hidden, `${file} : Row ${readRowNum}`).to.equal(
              readRowNum === 2
            );
            readRowNum++;
          }
          //  Check columns
          expect(
            worksheetReader.getColumn(1).hidden,
            `${file} : Column 1`
          ).to.equal(false);
          expect(
            worksheetReader.getColumn(2).hidden,
            `${file} : Column 2`
          ).to.equal(true);
          expect(
            worksheetReader.getColumn(3).hidden,
            `${file} : Column 3`
          ).to.equal(false);
        }
      });
    }
  });
});
