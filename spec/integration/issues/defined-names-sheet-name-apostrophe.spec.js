const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

// Print_Titles / Print_Area defined names embed the sheet name in formula syntax, where an
// apostrophe inside the quoted sheet name must be doubled (ECMA-376 formula grammar). Writing
// the raw sheet name produces an invalid defined name ('D'Alsace'!$1:$1) and Excel reports the
// workbook as corrupt ("We found a problem with some content...").
describe('github issues', () => {
  it('print titles and print areas escape apostrophes in sheet names', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('D\'Alsace');
    ws.getCell('A1').value = 'title row';
    ws.getCell('A2').value = 42;
    ws.pageSetup.printTitlesRow = '1:1';
    ws.pageSetup.printArea = 'A1:B2';

    const buffer = await wb.xlsx.writeBuffer();
    const zip = await JSZip.loadAsync(buffer);
    const workbookXml = await zip.file('xl/workbook.xml').async('string');
    const definedNames = workbookXml.match(/<definedName[^>]*>[^<]*<\/definedName>/g);
    expect(definedNames).to.have.length(2);
    definedNames.forEach(definedName => {
      expect(definedName).to.contain('&apos;D&apos;&apos;Alsace&apos;!');
    });
  });
});
