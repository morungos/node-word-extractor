gulp = require 'gulp'

mocha = require 'gulp-mocha'
coffee = require 'gulp-coffee'

gulp.task 'coffee', () ->
  gulp.src(['./src/**/*.*coffee'])
    .pipe coffee({bare: true})
    .pipe gulp.dest './lib/'

gulp.task 'test', ['coffee'], () ->
  gulp.src './test/**/*_test.coffee'
    .pipe mocha({reporter: 'spec'})

gulp.task 'build', ['coffee']
