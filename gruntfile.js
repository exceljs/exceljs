'use strict';

module.exports = function(grunt) {
  grunt.loadNpmTasks('grunt-babel');
  grunt.loadNpmTasks('grunt-browserify');
  grunt.loadNpmTasks('grunt-contrib-uglify');
  grunt.loadNpmTasks('grunt-contrib-jasmine');

  grunt.initConfig({
    babel: {
      options: {
        sourceMap: true
      },
      dist: {
        files: [
          {
            expand: true,
            src: ['./lib/**/*.js', './spec/browser/*.js'],
            dest: './build/'
          }
        ]
      }
    },
    browserify: {
      bundle: {
        src: ['./build/lib/exceljs.browser.js'],
        dest: './dist/exceljs.js',
        options: {
          browserifyOptions: {
            standalone: 'ExcelJS'
          }
        }
      },
      spec: {
        src: ['./build/spec/browser/exceljs.spec.js'],
        dest: './build/web/exceljs.spec.js'
      },
    },
    uglify: {
      options: {
        banner: '/*! ExcelJS <%= grunt.template.today("dd-mm-yyyy") %> */\n'
      },
      dist: {
        files: {
          './dist/exceljs.min.js': ['./dist/exceljs.js']
        }
      },
      es3: {
        files: [
          {
            expand: true,
            cwd: './build/lib/',
            src: ['*.js', '**/*.js'],
            dest: 'dist/es3/',
            ext: '.js',
          },
          {
            './dist/es3/index.js': ['./build/lib/exceljs.nodejs.js'],
          }
        ],
      },
    },

    jasmine: {
      dev: {
        src: ['./dist/exceljs.js'],
        options: {
          specs: './build/web/exceljs.spec.js'
        }
      }
    },
  });

  grunt.registerTask('build', ['babel', 'browserify', 'uglify']);
  grunt.registerTask('ug', ['uglify']);
};
