const fs = require('fs');

const exists = path => fs.promises.access(path).then(() => true).catch(() => false);

describe('ExcelJS', () => {
  describe('dist folder', () => {
    it('should include LICENSE', async () => {
      expect(await exists('./dist/LICENSE')).to.be.true();
    });
    it('should include dx-exceljs-fork.js', async () => {
      expect(await exists('./dist/dx-exceljs-fork.js')).to.be.true();
    });
    it('should include dx-exceljs-fork.min.js', async () => {
      expect(await exists('./dist/dx-exceljs-fork.min.js')).to.be.true();
    });
    it('should include dx-exceljs-fork.bare.js', async () => {
      expect(await exists('./dist/dx-exceljs-fork.bare.js')).to.be.true();
    });
    it('should include dx-exceljs-fork.bare.min.js', async () => {
      expect(await exists('./dist/dx-exceljs-fork.bare.min.js')).to.be.true();
    });
    it('should include es5/index', async () => {
      expect(await exists('./dist/es5/index.js')).to.be.true();
    });
  });
});
