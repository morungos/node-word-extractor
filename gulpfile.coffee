gulp = require 'gulp'
gulpif = require 'gulp-if'

mocha = require 'gulp-mocha'
coffee = require 'gulp-coffee'

gulp.task 'translate', () ->
  gulp.src(['./src/**/*.*'])
    .pipe gulpif /[.]coffee$/, coffee({bare: true})
    .pipe gulp.dest './lib/'

gulp.task 'test', ['translate'], () ->
  gulp.src './test/**/*_test.coffee'
    .pipe mocha({reporter: 'spec'})

gulp.task 'build', ['translate']
