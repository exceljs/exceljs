/* eslint-disable no-console */

const ExcelJS = verquire('exceljs');

const FORMULA_FILE_NAME = './spec/out/wb.formula.xlsx';
const FORMULA_ESCAPED_FILE_NAME = './spec/out/wb.formula.escaped.xlsx';

describe('WorkbookWriter', () => {
  describe('Serialise', () => {
    it('shared formula', () => {
      const options = {
        filename: FORMULA_FILE_NAME,
        useStyles: true,
      };
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter(options);
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = {
        formula: 'ROW()+COLUMN()',
        ref: 'A1:B2',
        result: 2,
      };
      ws.getCell('B1').value = {sharedFormula: 'A1', result: 3};
      ws.getCell('A2').value = {sharedFormula: 'A1', result: 3};
      ws.getCell('B2').value = {sharedFormula: 'A1', result: 4};

      ws.commit();
      return wb
        .commit()
        .then(() => {
          const wb2 = new ExcelJS.Workbook();
          return wb2.xlsx.readFile(FORMULA_FILE_NAME);
        })
        .then(wb2 => {
          const ws2 = wb2.getWorksheet('Hello');
          expect(ws2.getCell('A1').value).to.deep.equal({
            formula: 'ROW()+COLUMN()',
            shareType: 'shared',
            ref: 'A1:B2',
            result: 2,
          });
          expect(ws2.getCell('B1').value).to.deep.equal({
            sharedFormula: 'A1',
            result: 3,
          });
          expect(ws2.getCell('A2').value).to.deep.equal({
            sharedFormula: 'A1',
            result: 3,
          });
          expect(ws2.getCell('B2').value).to.deep.equal({
            sharedFormula: 'A1',
            result: 4,
          });
        });
    });

    it('shared formula escaped', () => {
      const options = {
        filename: FORMULA_ESCAPED_FILE_NAME,
        useStyles: true,
      };
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter(options);
      const ws = wb.addWorksheet('Hello');
      ws.getCell('A1').value = '=RIF.RIGA()+RIF.COLONNA()';
      ws.getCell('A1').quotePrefix = true;
      ws.getCell('B1').value = '=RIF.RIGA()+RIF.COLONNA()';
      ws.getCell('A2').value = '=RIF.RIGA()+RIF.COLONNA()';
      ws.getCell('B2').value = '=RIF.RIGA()+RIF.COLONNA()';

      ws.commit();
      return wb.commit().then(() => {
        const wb2 = new ExcelJS.Workbook();
        return wb2.xlsx.readFile(FORMULA_ESCAPED_FILE_NAME);
      });
    });
  });
});
