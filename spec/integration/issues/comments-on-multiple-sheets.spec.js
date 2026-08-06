const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

// Excel allocates VML comment shape ids workbook-globally in blocks of 1024 per drawing part
// (ISO 29500 §14.2.2.14 "idmap") and expects each part to reserve a distinct block. Reusing
// block "1" and shape ids _x0000_s1025... in every sheet's vmlDrawingN.vml makes Excel report
// the file as corrupt ("We found a problem with some content...") as soon as two or more
// sheets have comments.
describe('github issues', () => {
  it('comments on multiple sheets produce workbook-unique VML shape ids', async () => {
    const wb = new ExcelJS.Workbook();
    ['One', 'Two', 'Three'].forEach(name => {
      const ws = wb.addWorksheet(name);
      ws.getCell('A1').value = name;
      ws.getCell('A1').note = `first note on ${name}`;
      ws.getCell('C3').value = 42;
      ws.getCell('C3').note = `second note on ${name}`;
    });

    const buffer = await wb.xlsx.writeBuffer();
    const zip = await JSZip.loadAsync(buffer);
    const vmlNames = Object.keys(zip.files).filter(name =>
      /^xl\/drawings\/vmlDrawing\d+\.vml$/.test(name)
    );
    expect(vmlNames.length).to.equal(3);

    const shapeIds = [];
    const idmapBlocks = [];
    for (const name of vmlNames) {
      // eslint-disable-next-line no-await-in-loop
      const xml = await zip.file(name).async('string');
      shapeIds.push(...xml.match(/id="_x0000_s\d+"/g));
      idmapBlocks.push(...xml.match(/<o:idmap [^>]*data="\d+"/g));
    }
    expect(shapeIds.length).to.equal(6);
    expect(new Set(shapeIds).size).to.equal(shapeIds.length);
    expect(new Set(idmapBlocks).size).to.equal(idmapBlocks.length);
  });
});
