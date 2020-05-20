import 'regenerator-runtime/runtime';

import { expect } from 'chai';
import ExcelJS from '../../index';

describe('typescript', () => {
  it('can create and buffer xlsx', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = 7;
    const buffer = await wb.xlsx.writeBuffer({
      useStyles: true,
      useSharedStrings: true,
    });

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet('blort');
    expect(ws2.getCell('A1').value).to.equal(7);
  });
  it('can create and stream xlsx', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = 7;

    const wb2 = new ExcelJS.Workbook();
    const stream = wb2.xlsx.createInputStream();
    await wb.xlsx.write(stream);
    stream.end();

    await new Promise((resolve, reject) => {
      stream.on('done', () => {
        const ws2 = wb2.getWorksheet('blort');
        expect(ws2.getCell('A1').value).to.equal(7);
        resolve();
      });
      stream.on('error', reject);
    })
  });
});
