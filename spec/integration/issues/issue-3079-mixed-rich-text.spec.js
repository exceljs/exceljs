const ExcelJS = verquire('exceljs');
const JSZip = require('jszip');

describe('github issues', () => {
  it('issue 3079 - loads a string that has plain text and rich-text runs', async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('S');
    worksheet.getCell('A1').value = 'placeholder';
    worksheet.getCell('B1').value = 'empty-leading';

    const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
    const sharedStrings = (await zip.file('xl/sharedStrings.xml').async('string'))
      .replace(
        '<si><t>placeholder</t></si>',
        '<si><t>xy</t><r><rPr><u val="single"/><color theme="10"/></rPr><t>link text</t></r></si>'
      )
      .replace(
        '<si><t>empty-leading</t></si>',
        '<si><t></t><r><t>only run</t></r></si>'
      );
    zip.file('xl/sharedStrings.xml', sharedStrings);

    const sheet = (await zip.file('xl/worksheets/sheet1.xml').async('string')).replace(
      '</row>',
      '<c r="C1" t="inlineStr"><is><t>pre</t><r><t>post</t></r></is></c></row>'
    );
    zip.file('xl/worksheets/sheet1.xml', sheet);

    const loaded = new ExcelJS.Workbook();
    await loaded.xlsx.load(await zip.generateAsync({type: 'nodebuffer'}));
    const sheetOut = loaded.getWorksheet('S');

    expect(sheetOut.getCell('A1').text).to.equal('xylink text');
    expect(sheetOut.getCell('A1').value).to.deep.equal({
      richText: [
        {text: 'xy'},
        {font: {underline: 'single', color: {theme: 10}}, text: 'link text'},
      ],
    });
    expect(sheetOut.getCell('B1').text).to.equal('only run');
    expect(sheetOut.getCell('C1').text).to.equal('prepost');
    expect(sheetOut.getCell('C1').value).to.deep.equal({
      richText: [{text: 'pre'}, {text: 'post'}],
    });
  });
});
