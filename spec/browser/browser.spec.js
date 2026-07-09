const fs = require("fs");
const path = require("path");
const { JSDOM } = require("jsdom");

// The browser bundle (dist/exceljs.js) is a UMD build produced by `npm run build`.
// Loading it inside a jsdom window exercises the browserify shims (crypto-browserify,
// Buffer, etc.) exactly as a real browser would, without requiring a headless
// browser binary (puppeteer/playwright) or grunt-contrib-jasmine.

describe("browser bundle (dist/exceljs.js loaded in jsdom)", function () {
  this.timeout(30000);

  let ExcelJS;

  before(() => {
    const bundlePath = path.join(__dirname, "..", "..", "dist", "exceljs.js");
    if (!fs.existsSync(bundlePath)) {
      throw new Error(
        "dist/exceljs.js not found. Run `npm run build` before the browser test.",
      );
    }
    const dom = new JSDOM("<!DOCTYPE html>", { runScripts: "outside-only" });
    const { window } = dom;
    // The UMD bundle assigns to `window.ExcelJS` when module/exports/define are absent.
    window.eval(fs.readFileSync(bundlePath, "utf8"));
    ExcelJS = window.ExcelJS;
    if (!ExcelJS || !ExcelJS.Workbook) {
      throw new Error("Bundle did not expose window.ExcelJS.Workbook");
    }
  });

  it("should read and write xlsx via binary buffer", async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("blort");

    ws.getCell("A1").value = "Hello, World!";
    ws.getCell("A2").value = 7;

    const buffer = await wb.xlsx.writeBuffer();
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer);
    const ws2 = wb2.getWorksheet("blort");
    expect(ws2).to.be.ok;
    expect(ws2.getCell("A1").value).to.equal("Hello, World!");
    expect(ws2.getCell("A2").value).to.equal(7);
  });

  it("should read and write xlsx via base64 buffer", async () => {
    const options = { base64: true };
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("blort");

    ws.getCell("A1").value = "Hello, World!";
    ws.getCell("A2").value = 7;

    const buffer = await wb.xlsx.writeBuffer(options);
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buffer.toString("base64"), options);
    const ws2 = wb2.getWorksheet("blort");
    expect(ws2).to.be.ok;
    expect(ws2.getCell("A1").value).to.equal("Hello, World!");
    expect(ws2.getCell("A2").value).to.equal(7);
  });

  it("should write csv via buffer", async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("blort");

    ws.getCell("A1").value = "Hello, World!";
    ws.getCell("B1").value = "What time is it?";
    ws.getCell("A2").value = 7;
    ws.getCell("B2").value = "12pm";

    const buffer = await wb.csv.writeBuffer();
    expect(buffer.toString()).to.equal(
      '"Hello, World!",What time is it?\n7,12pm',
    );
  });
});
