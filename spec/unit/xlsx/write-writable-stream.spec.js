const ExcelJS = verquire('exceljs');

describe('Web-native streams', () => {
  // eslint-disable-next-line no-use-before-define
  const {TransformStream} = global;
  const skip = typeof TransformStream !== 'function';

  async function testTransformStream(transform) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('blort');
    ws.getCell('A1').value = 7;

    const wb2 = new ExcelJS.Workbook();

    await Promise.all([
      wb.xlsx.write(transform.writable),
      wb2.xlsx.read(transform.readable),
    ]);

    const ws2 = wb2.getWorksheet('blort');
    expect(ws2.getCell('A1').value).to.equal(7);
  }

  it('can stream xlsx through a web-native TransformStream (readable/writable pair)', async function() {
    if (skip) {
      this.skip();
    }

    await testTransformStream(new TransformStream());
  });

  it('handles ReadableStream lacking @@asyncIterator (Safari compat)', async function() {
    if (skip) {
      this.skip();
    }

    const safariTransformStream = new TransformStream();

    Object.defineProperty(
      safariTransformStream.readable,
      Symbol.asyncIterator,
      {
        value: undefined,
      }
    );

    await testTransformStream(safariTransformStream);
  });
});
