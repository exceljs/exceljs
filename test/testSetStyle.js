const Excel = require('../lib/exceljs.nodejs.js');

const wb = new Excel.Workbook();
wb.xlsx
  .readFile(require.resolve('./data/testStyle.xlsx'))
  .then(() => {
    wb.worksheets.forEach(sheet => {
      // May be different ?
      console.info(
        `Compare A1 to A2 : ${
          sheet.getCell('A1').style === sheet.getCell('A2').style}`
      );
      console.info(
        `Compare F1 to J1 : ${
          sheet.getCell('F1').style === sheet.getCell('J1').style}`
      );

      // supposed to be equal
      console.info(
        `Compare B2 to D5 : ${
          sheet.getCell('B2').style === sheet.getCell('D5').style}`
      );
      console.info(
        `Compare G2 to I5 : ${
          sheet.getCell('G2').style === sheet.getCell('I5').style}`
      );
      console.info(
        `Compare E11 to G13 : ${
          sheet.getCell('E11').style === sheet.getCell('G13').style}`
      );

      // supposed to be different
      console.info(
        `Compare B2 to G13 : ${
          sheet.getCell('B2').style === sheet.getCell('G13').style}`
      );
      console.info(
        `Compare B2 to G2 : ${
          sheet.getCell('B2').style === sheet.getCell('G2').style}`
      );
      console.info(
        `Compare G2 to G13 : ${
          sheet.getCell('G2').style === sheet.getCell('G13').style}`
      );

      // supposed to change B2 to D9 font color
      const b2 = sheet.getCell('B2');
      b2.font = {...b2.font, color: {argb: 'FF123456'}};

      // supposed to change only C12 font color
      const c12 = sheet.getCell('C12');
      c12.setStyle({font: {...b2.font, color: {argb: 'FF123456'}}});

      // supposed to change G2 to I9 background color
      sheet.getCell('G2').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: 'FF789ABC'},
      };

      // supposed to change only F13 background color
      sheet.getCell('F13').setStyle({
        fill: {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {argb: 'FF789ABC'},
        },
      });
    });

    return wb.xlsx.writeFile(`${__dirname}/data/testStyleResult.xlsx`);
  })
  .catch(console.error);
