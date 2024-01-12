const ExcelJS = verquire('exceljs');
const {createReadStream} = require('fs');

const fileName = './spec/integration/data/third-party-export.xlsx';

const cellValues = {
  A1: 'Record ID',
  B1: 'Column 1',
  C1: 'Column 2',
  D1: 'Date',
  E1: 'Column 3',
  A2: '8209375239',
  B2: 'Data 1',
  C2: 'Data 2',
  // D2: '2022-03-17T15:10:00.000Z', Don't compare date value
  E2: 'Data 3',
};

describe('github issues', () => {
  describe('third party export excel file', () => {
    it('reads correctly when using readFile', async () => {
      const wb = new ExcelJS.Workbook();

      await wb.xlsx.readFile(fileName);

      const worksheet = wb.getWorksheet(1);
      expect(worksheet.rowCount).to.equal(2);
      expect(worksheet.actualRowCount).to.equal(2);

      for (const [cell, value] of Object.entries(cellValues)) {
        expect(worksheet.getCell(cell).value).to.equal(value);
      }
    });

    it('reads correctly when reading from a stream', async () => {
      const fileStream = createReadStream(fileName);
      const wb = new ExcelJS.stream.xlsx.WorkbookReader(fileStream);

      for await (const worksheet of wb) {
        let rowCount = 0;
        for await (const row of worksheet) {
          rowCount += 1;
          let cellCount = 0;

          row.eachCell(cell => {
            cellCount += 1;
            if (cell._address === 'D2') return; // Don't compare date value
            expect(cell.value).to.equal(cellValues[cell._address]);
          });
          expect(cellCount).to.equal(5);
        }
        expect(rowCount).to.equal(2);

        break; // Only first worksheet
      }
    });
  });
});
