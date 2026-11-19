const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

describe('github issues', () => {
  it('issue 3115 - reads a hyperlink shared by a range of cells', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Links');
    const target = 'https://example.com/shared';
    sheet.getCell('A1').value = {text: 'single', hyperlink: 'https://example.com/single'};
    sheet.getCell('B2').value = {text: 'first', hyperlink: target};
    sheet.getCell('C2').value = 'second';
    sheet.getCell('B4').value = 'third';
    sheet.getCell('C4').value = 'last';
    sheet.getCell('D4').value = 'outside';

    // Excel can write a single hyperlink ref for an entire pasted rectangle.
    const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
    const path = 'xl/worksheets/sheet1.xml';
    const xml = await zip.file(path).async('string');
    zip.file(path, xml.replace('<hyperlink ref="B2"', '<hyperlink ref="B2:C4"'));

    const loaded = new ExcelJS.Workbook();
    await loaded.xlsx.load(await zip.generateAsync({type: 'nodebuffer'}));
    const links = loaded.getWorksheet('Links');
    expect(links.getCell('B2').value).to.deep.equal({text: 'first', hyperlink: target});
    expect(links.getCell('C2').value).to.deep.equal({text: 'second', hyperlink: target});
    expect(links.getCell('B4').value).to.deep.equal({text: 'third', hyperlink: target});
    expect(links.getCell('C4').value).to.deep.equal({text: 'last', hyperlink: target});
    expect(links.getCell('A1').hyperlink).to.equal('https://example.com/single');
    expect(links.getCell('D4').value).to.equal('outside');
  });
});
