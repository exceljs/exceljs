'use strict';

module.exports = function(grunt) {
  grunt.loadNpmTasks('grunt-babel');
  grunt.loadNpmTasks('grunt-browserify');
  grunt.loadNpmTasks('grunt-terser');
  grunt.loadNpmTasks('grunt-contrib-jasmine');
  grunt.loadNpmTasks('grunt-contrib-copy');
  grunt.loadNpmTasks('grunt-exorcise');

  grunt.initConfig({
    babel: {
      options: {
        sourceMap: true,
        compact: false,
        targets: {
          node:"current"
        }
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
        transform: [
          [
            'babelify',
            {
              // enable babel transpile for node_modules
              global: true,
              presets: ['@babel/preset-env'],
              // core-js should not be transpiled
              // See https://github.com/zloirock/core-js/issues/514
              ignore: [/node_modules[\\/]core-js/],
            },
          ],
        ],
        browserifyOptions: {
          // enable source map for browserify
          debug: true,
          standalone: 'ExcelJS',
        },
      },
      bundle: {
        // keep the original source for source maps
        src: ['./lib/exceljs.browser.js'],
        dest: './dist/exceljs.js',
      },
    },

    terser: {
      options: {
        output: {
          preamble: '/*! ExcelJS <%= grunt.template.today("dd-mm-yyyy") %> */\n',
          ascii_only: true,
        },
      },
      dist: {
        options: {
          // Keep the original source maps from browserify
          // See also https://www.npmjs.com/package/terser#source-map-options
          sourceMap: {
            content: 'inline',
            url: 'exceljs.min.js.map',
          },
        },
        files: {
          './dist/exceljs.min.js': ['./dist/exceljs.js'],
        },
      },
    },

    // Move source maps to a separate file
    exorcise: {
      bundle: {
        options: {},
        files: {
          './dist/exceljs.js.map': ['./dist/exceljs.js'],
        },
      },
    },

    copy: {
      dist: {
        files: [
          {expand: true, src: ['**'], cwd: './build/lib', dest: './dist/es5'},
          {src: './build/lib/exceljs.nodejs.js', dest: './dist/es5/index.js'},
          {src: './LICENSE', dest: './dist/LICENSE'},
        ],
      },
    },

  });

  grunt.registerTask('build', ['babel:dist', 'browserify', 'terser', 'exorcise', 'copy']);
  grunt.registerTask('ug', ['terser']);
};
