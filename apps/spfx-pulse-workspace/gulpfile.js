'use strict';

const build = require('@microsoft/sp-build-web');
const gulp = require('gulp');
const path = require('path');

build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    generatedConfiguration.resolve = generatedConfiguration.resolve || {};
    generatedConfiguration.resolve.alias = {
      ...(generatedConfiguration.resolve.alias || {}),
      '@': path.resolve(__dirname, 'lib')
    };

    return generatedConfiguration;
  }
});

build.initialize(gulp);

gulp.task('serve', gulp.series('serve-deprecated'));
