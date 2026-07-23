'use strict';

const path = require('path');
const {Transform} = require('stream');
// eslint-disable-next-line import/no-extraneous-dependencies
const babel = require('@babel/core');

const pkg = require('./package.json');

// Browserify transform that transpiles sources and dependencies with Babel.
// Replaces the unmaintained `babelify` package, which is incompatible with
// Babel 8 (it calls the now callback-only `loadPartialConfig` synchronously).
function babelTransform(filename) {
  if (!babel.DEFAULT_EXTENSIONS.includes(path.extname(filename))) {
    return new Transform({transform: (chunk, enc, cb) => cb(null, chunk)});
  }

  const chunks = [];
  return new Transform({
    transform(chunk, enc, cb) {
      chunks.push(chunk);
      cb();
    },
    flush(cb) {
      const code = Buffer.concat(chunks).toString();
      babel.transform(
        code,
        {
          filename,
          presets: ['@babel/preset-env'],
          // core-js should not be transpiled
          // See https://github.com/zloirock/core-js/issues/514
          ignore: [/node_modules[\\/]core-js/],
          sourceMaps: 'inline',
        },
        (err, result) => {
          if (err) {
            cb(err);
            return;
          }
          cb(null, result ? result.code : code);
        }
      );
    },
  });
}

/* eslint-disable max-len */
const preamble = `/*!
 * DevExtreme-ExcelJS Fork v.${pkg.version}
 * https://js.devexpress.com/
 * Copyright (c) 2025, Developer Express Inc.
 * Copyright (c) 2014-2019 Guyon Roche
 * Read about DevExtreme-ExcelJS Fork licensing here: https://cdn.jsdelivr.net/npm/devextreme-exceljs-fork@${pkg.version}/LICENSE
 */`;
/* eslint-enable max-len */

module.exports = function(grunt) {
  grunt.loadNpmTasks('grunt-babel');
  grunt.loadNpmTasks('grunt-browserify');
  grunt.loadNpmTasks('grunt-terser');
  grunt.loadNpmTasks('grunt-contrib-copy');
  grunt.loadNpmTasks('grunt-exorcise');

  grunt.initConfig({
    babel: {
      options: {
        sourceMap: true,
        compact: false,
      },
      dist: {
        files: [
          {
            expand: true,
            src: ['./lib/**/*.js'],
            dest: './build/',
          },
        ],
      },
    },
    browserify: {
      options: {
        banner: preamble,
        transform: [
          // global: true enables transpile for node_modules too
          [babelTransform, {global: true}],
        ],
        browserifyOptions: {
          // enable source map for browserify
          debug: true,
          standalone: 'ExcelJS',
        },
      },
      bare: {
        // Entries are specified explicitly via browserifyOptions to work around
        // glob v13 + minimatch v10 treating backslashes as escape chars on Windows,
        // which causes grunt.file.expand to return empty results.
        options: {
          browserifyOptions: {
            debug: true,
            standalone: 'ExcelJS',
            entries: ['./lib/dx-exceljs-fork.bare.js'],
          },
        },
        src: [],
        dest: './dist/dx-exceljs-fork.bare.js',
      },
      bundle: {
        options: {
          browserifyOptions: {
            debug: true,
            standalone: 'ExcelJS',
            entries: ['./lib/dx-exceljs-fork.browser.js'],
          },
        },
        src: [],
        dest: './dist/dx-exceljs-fork.js',
      },
    },

    terser: {
      options: {
        output: {
          ascii_only: true,
        },
      },
      dist: {
        options: {
          // Keep the original source maps from browserify
          // See also https://www.npmjs.com/package/terser#source-map-options
          sourceMap: {
            content: 'inline',
            url: 'dx-exceljs-fork.min.js.map',
          },
        },
        files: {
          './dist/dx-exceljs-fork.min.js': ['./dist/dx-exceljs-fork.js'],
        },
      },
      bare: {
        options: {
          // Keep the original source maps from browserify
          // See also https://www.npmjs.com/package/terser#source-map-options
          sourceMap: {
            content: 'inline',
            url: 'dx-exceljs-fork.bare.min.js.map',
          },
        },
        files: {
          './dist/dx-exceljs-fork.bare.min.js': ['./dist/dx-exceljs-fork.bare.js'],
        },
      },
    },

    // Move source maps to a separate file
    exorcise: {
      bundle: {
        options: {},
        files: {
          './dist/dx-exceljs-fork.js.map': ['./dist/dx-exceljs-fork.js'],
          './dist/dx-exceljs-fork.bare.js.map': ['./dist/dx-exceljs-fork.bare.js'],
        },
      },
    },

    copy: {
      dist: {
        files: [
          {
            expand: true, src: ['**'], cwd: './build/lib', dest: './dist/es5',
          },
          {src: './build/lib/dx-exceljs-fork.nodejs.js', dest: './dist/es5/index.js'},
          {src: './LICENSE', dest: './dist/LICENSE'},
        ],
      },
    },
  });

  grunt.registerTask('build', ['babel:dist', 'browserify', 'terser', 'exorcise', 'copy']);
  grunt.registerTask('ug', ['terser']);
};
