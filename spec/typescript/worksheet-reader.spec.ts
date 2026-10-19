import 'regenerator-runtime/runtime';

import {expect} from 'chai';
import {Readable} from 'stream';
// Import the published declarations explicitly; index.ts wraps untyped ES5 output.
import type * as ExcelTypes from '../../index.d';

const ExcelJS: typeof ExcelTypes = require('../../excel');

for (const useSharedStrings of [false, true]) {
  describe(`streaming worksheet names (shared strings: ${useSharedStrings})`, () => {
    it('identifies each worksheet while consuming its rows', async () => {
      const workbook = new ExcelJS.Workbook();
      const names = ['Quarterly Sales', '支出 & costs'];
      names.forEach(name => {
        workbook.addWorksheet(name).addRow([name]);
      });
      const buffer = await workbook.xlsx.writeBuffer({useSharedStrings});
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(
        Readable.from([buffer]),
        {}
      );
      const actual: string[] = [];
      for await (const worksheet of reader) {
        const name: string = worksheet.name;
        actual.push(name);
        const values: ExcelTypes.CellValue[] = [];
        for await (const row of worksheet) {
          values.push(row.getCell(1).value);
        }
        expect(values).to.deep.equal([name]);
      }
      expect(actual).to.deep.equal(names);
    });
  });
}
