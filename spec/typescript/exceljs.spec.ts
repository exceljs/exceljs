import 'regenerator-runtime/runtime';

import { expect } from 'chai';
import _ExcelJS from '../../index';
import type * as ExcelJSModule from '../../exceljs';

const ExcelJS = _ExcelJS as unknown as typeof ExcelJSModule;

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
    const ws2 = wb2.getWorksheet('blort')!;
    expect(ws2.getCell('A1').value).to.equal(7);
  });
  it('can create and stream xlsx', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = 7;

    const wb2 = new ExcelJS.Workbook();
    const transform = new TransformStream<Uint8Array, Uint8Array>();

    // type checking - code should compile due to `@ts-expect-error`
    // but will not run as function is never invoked
    void (() => {
      // @ts-expect-error
      wb.xlsx.write({});
    });

    await Promise.all([
      wb.xlsx.write(transform.writable),
      wb2.xlsx.read(transform.readable),
    ]);

    const ws2 = wb2.getWorksheet('blort')!;
    expect(ws2.getCell('A1').value).to.equal(7);
  });
});
