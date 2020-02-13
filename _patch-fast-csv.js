const fs = require('fs');

const fastCsvPackage = JSON.parse(fs.readFileSync(`${__dirname}/node_modules/fast-csv/package.json`));

fastCsvPackage.browserify = {
  transform: [['babelify', {presets: ['@babel/preset-env']}]],
};

fs.writeFileSync(`${__dirname}/node_modules/fast-csv/package.json`, JSON.stringify(fastCsvPackage, null, 2));
