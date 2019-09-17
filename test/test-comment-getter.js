const {Workbook} = require('../lib/exceljs.nodejs');

( async ()=> {
    try {
        const wb = new Workbook();
        await wb.xlsx.readFile(require.resolve('./data/comments.xlsx'));
        wb.eachSheet(sheet => {
            sheet.eachRow(row => {
                row.eachCell(cell => {
                    console.info(cell.comment);
                });
            });
        });
    } catch(error) {
        throw error;
    }
})();