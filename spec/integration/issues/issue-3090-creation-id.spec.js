const ExcelJS = verquire('exceljs');
const JSZip = require('jszip');

const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';

describe('github issues', () => {
  it('issue 3090 - gives each picture its own creationId', async () => {
    const workbook = new ExcelJS.Workbook();
    const images = workbook.addWorksheet('Images');
    const red = workbook.addImage({base64: PNG, extension: 'png'});
    const blue = workbook.addImage({base64: PNG, extension: 'png'});
    images.addImage(red, {tl: {col: 0, row: 0}, ext: {width: 100, height: 100}});
    images.addImage(blue, {tl: {col: 0, row: 10}, ext: {width: 100, height: 100}});

    // A worksheet with no pictures is what makes Excel for the web drop
    // every picture that shares one creationId.
    workbook.addWorksheet('No images');

    const more = workbook.addWorksheet('More');
    const green = workbook.addImage({base64: PNG, extension: 'png'});
    more.addImage(green, {tl: {col: 1, row: 1}, ext: {width: 50, height: 50}});

    const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer());
    const drawingNames = Object.keys(zip.files)
      .filter(name => /^xl\/drawings\/drawing\d+\.xml$/.test(name))
      .sort();
    expect(drawingNames).to.deep.equal(['xl/drawings/drawing1.xml', 'xl/drawings/drawing2.xml']);

    const xmls = await Promise.all(drawingNames.map(name => zip.file(name).async('string')));
    const ids = [];
    xmls.forEach(xml => {
      const found = xml.match(/<a16:creationId\b[^>]*\bid="(\{[^"]+\})"/g) || [];
      found.forEach(tag => {
        ids.push(tag.replace(/.*\bid="(\{[^"]+\})".*/, '$1'));
      });
    });

    expect(ids).to.deep.equal([
      '{00000000-0008-0000-0000-000000000001}',
      '{00000000-0008-0000-0000-000000000002}',
      '{00000000-0008-0000-0000-000000000003}',
    ]);
  });
});
