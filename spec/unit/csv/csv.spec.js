const ExcelJS = verquire('exceljs');

describe('CSV security', () => {
  it('escapes spreadsheet formula values by default', async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet 1');
    worksheet.addRow(['=1+1', '+1', '-1', '@cmd', '\tvalue', '\rvalue', 'safe']);

    const csv = (await workbook.csv.writeBuffer()).toString();

    expect(csv).to.contain('\'=1+1');
    expect(csv).to.contain('\'+1');
    expect(csv).to.contain('\'-1');
    expect(csv).to.contain('\'@cmd');
    expect(csv).to.contain('\'\tvalue');
    expect(csv).to.contain('\'\rvalue');
    expect(csv).to.contain('safe');
  });
});
