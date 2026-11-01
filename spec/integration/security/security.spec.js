const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

async function expectLoadToFail(workbook, buffer, options, message) {
  let error;
  try {
    await workbook.xlsx.load(buffer, options);
  } catch (caught) {
    error = caught;
  }
  expect(error).to.be.an('error');
  expect(error.message).to.contain(message);
}

describe('Security hardening', () => {
  afterEach(() => {
    delete Object.prototype.exceljsPolluted;
  });

  it('rejects image filenames containing traversal segments', () => {
    const workbook = new ExcelJS.Workbook();

    expect(() =>
      workbook.addImage({filename: '../secret.txt', extension: 'png'})
    ).to.throw(/path traversal/);
  });

  it('rejects oversized ZIP entries before decompression', async () => {
    const zip = new JSZip();
    zip.file('payload.bin', 'payload');
    const buffer = await zip.generateAsync({type: 'nodebuffer'});

    await expectLoadToFail(
      new ExcelJS.Workbook(),
      buffer,
      {maxEntrySize: 1},
      'maxEntrySize'
    );
  });

  it('rejects ZIP compression bombs by ratio', async () => {
    const zip = new JSZip();
    zip.file('payload.bin', 'A'.repeat(10000));
    const buffer = await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: {level: 9},
    });

    await expectLoadToFail(
      new ExcelJS.Workbook(),
      buffer,
      {maxCompressionRatio: 2},
      'maxCompressionRatio'
    );
  });

  it('limits streamed ZIP entries with data descriptors', async () => {
    const reader = new ExcelJS.stream.xlsx.WorkbookReader('spec/integration/data/gold.xlsx', {
      maxEntrySize: 1,
    });
    const error = await new Promise(resolve => {
      reader.once('error', resolve);
      reader.read();
    });

    expect(error).to.be.an('error');
    expect(error.message).to.contain('maxEntrySize');
  });
});
