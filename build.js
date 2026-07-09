"use strict";

/* eslint-disable no-console */

const fs = require("fs");
const path = require("path");
const { promisify } = require("util");
const { glob } = require("glob");
const babel = require("@babel/core");
const browserify = require("browserify");
const terser = require("terser");
const exorcist = require("exorcist");

const ROOT = __dirname;
const readFile = promisify(fs.readFile);
const writeFile = promisify(fs.writeFile);
const mkdir = promisify(fs.mkdir);

async function ensureDir(dir) {
  await mkdir(dir, { recursive: true });
}

async function rmrf(dir) {
  await fs.promises.rm(dir, { recursive: true, force: true });
}

// 1. Babel: transpile ./lib/**/*.js and ./spec/browser/*.js into ./build/
async function babelBuild() {
  const libFiles = await glob("lib/**/*.js", { ignore: "node_modules/**" });
  const specFiles = await glob("spec/browser/*.js");
  const files = [...libFiles, ...specFiles];
  for (const file of files) {
    const result = await babel.transformFileAsync(path.join(ROOT, file), {
      sourceMap: true,
      compact: false,
    });
    const outPath = path.join(ROOT, "./build", file);
    await ensureDir(path.dirname(outPath));
    let output = result.code;
    if (result.map) {
      output += `\n//# sourceMappingURL=${path.basename(outPath)}.map`;
      await writeFile(`${outPath}.map`, JSON.stringify(result.map));
    }
    await writeFile(outPath, output);
  }
  console.log(`Babel: transpiled ${files.length} files -> build/`);
}

// 2. Browserify: bundle with babelify transform
function bundleEntry(src, dest, opts) {
  return new Promise((resolve, reject) => {
    const b = browserify(src, {
      debug: true,
      standalone: "ExcelJS",
      ...opts.browserifyOptions,
    });
    if (opts.transform) {
      b.transform(opts.transform[0][0], opts.transform[0][1]);
    }
    const stream = b.bundle();
    const out = fs.createWriteStream(path.join(ROOT, dest));
    stream.pipe(out);
    out.on("finish", resolve);
    stream.on("error", reject);
    out.on("error", reject);
  });
}

const babelifyTransform = [
  [
    "babelify",
    {
      global: true,
      presets: ["@babel/preset-env"],
      ignore: [/node_modules[\\/]core-js/],
    },
  ],
];

async function browserifyBuild() {
  await ensureDir(path.join(ROOT, "./dist"));

  // bundle: ./lib/exceljs.browser.js -> ./dist/exceljs.js
  await bundleEntry("./lib/exceljs.browser.js", "./dist/exceljs.js", {
    transform: babelifyTransform,
  });
  console.log("Bundle ./dist/exceljs.js created.");

  // bare: ./lib/exceljs.bare.js -> ./dist/exceljs.bare.js
  await bundleEntry("./lib/exceljs.bare.js", "./dist/exceljs.bare.js", {
    transform: babelifyTransform,
  });
  console.log("Bundle ./dist/exceljs.bare.js created.");
}

// 3. Terser: minify with source maps from inline
async function minifyFile(src, dest, mapUrl) {
  const code = await readFile(path.join(ROOT, src), "utf8");
  const result = await terser.minify(code, {
    sourceMap: {
      content: "inline",
      url: mapUrl,
    },
    output: {
      preamble: "/*! ExcelJS */\n",
      ascii_only: true,
    },
  });
  if (result.error) throw result.error;
  await writeFile(path.join(ROOT, dest), result.code);
  if (result.map) {
    await writeFile(path.join(ROOT, `${dest}.map`), result.map);
  }
  console.log(`Minified ${dest}`);
}

async function terserBuild() {
  await minifyFile(
    "./dist/exceljs.js",
    "./dist/exceljs.min.js",
    "exceljs.min.js.map",
  );
  await minifyFile(
    "./dist/exceljs.bare.js",
    "./dist/exceljs.bare.min.js",
    "exceljs.bare.min.js.map",
  );
}

// 4. Exorcise: extract inline source maps to separate files
function exorciseFile(src, mapDest) {
  return new Promise((resolve, reject) => {
    const srcPath = path.join(ROOT, src);
    const tmpPath = `${srcPath}.tmp`;
    fs.createReadStream(srcPath)
      .pipe(exorcist(path.join(ROOT, mapDest)))
      .pipe(fs.createWriteStream(tmpPath))
      .on("finish", () => {
        fs.renameSync(tmpPath, srcPath);
        resolve();
      })
      .on("error", reject);
  });
}

async function exorciseBuild() {
  await exorciseFile("./dist/exceljs.js", "./dist/exceljs.js.map");
  await exorciseFile("./dist/exceljs.bare.js", "./dist/exceljs.bare.js.map");
  console.log("Source maps extracted.");
}

// 5. Copy: build/lib -> dist/es5, index, LICENSE
async function copyBuild() {
  await rmrf(path.join(ROOT, "./dist/es5"));
  await fs.promises.cp(
    path.join(ROOT, "./build/lib"),
    path.join(ROOT, "./dist/es5"),
    {
      recursive: true,
    },
  );
  await fs.promises.copyFile(
    path.join(ROOT, "./build/lib/exceljs.nodejs.js"),
    path.join(ROOT, "./dist/es5/index.js"),
  );
  await fs.promises.copyFile(
    path.join(ROOT, "./LICENSE"),
    path.join(ROOT, "./dist/LICENSE"),
  );
  console.log("Copied to dist/es5.");
}

async function main() {
  await rmrf(path.join(ROOT, "./build"));
  await babelBuild();
  await browserifyBuild();
  await terserBuild();
  await exorciseBuild();
  await copyBuild();
  console.log("Build complete.");
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
