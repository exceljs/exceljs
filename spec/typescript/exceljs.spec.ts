import "regenerator-runtime/runtime";

import { expect } from "chai";
import * as ExcelJS from "../../index";
import { createReadStream, createWriteStream } from "node:fs";

describe("typescript", () => {
  it("can create and buffer xlsx", async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("blort");
    ws.getCell("A1").value = 7;
    const buffer = await wb.xlsx.writeBuffer({
      useStyles: true,
      useSharedStrings: true,
    });

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet("blort");
    expect(ws2.getCell("A1").value).to.equal(7);
  });

  it("can stream xlsx", async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("blort");
    ws.getCell("A1").value = 7;

    const write = createWriteStream("test.xlsx");

    await wb.xlsx.write(write);

    const read = createReadStream("test.xlsx");

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.read(read);
    const ws2 = wb2.getWorksheet("blort");

    expect(ws2.getCell("A1").value).to.equal(7);
  });
});
