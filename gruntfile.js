'use strict';

module.exports = function(grunt) {

  grunt.loadNpmTasks('grunt-babel');
  grunt.loadNpmTasks('grunt-browserify');
  grunt.loadNpmTasks('grunt-contrib-uglify');

  grunt.initConfig({
    babel: {
      options: {
        sourceMap: true
      },
      dist: {
        files: [
          {
            expand: true,
            src: ['./lib/**/*.js'],
            dest: './build/'
          }
        ]
      }
    },
    browserify: {
      standalone: {
        src: ['./build/lib/exceljs.browser.js'],
        dest: './dist/exceljs.js',
        options: {
          browserifyOptions: {
            standalone: 'ExcelJS'
          }
        }
      }
    },
    uglify: {
      options: {
        banner: '/*! ExcelJS <%= grunt.template.today("dd-mm-yyyy") %> */\n'
      },
      dist: {
        files: {
          './dist/exceljs.min.js': ['./dist/exceljs.js']
        }
      }
    }
  });

  grunt.registerTask('build', ['babel', 'browserify', 'uglify']);
};
