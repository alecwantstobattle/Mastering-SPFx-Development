'use strict';

if (process.argv.indexOf('all-in-one-go') !== -1) {
  process.argv.push('--ship');
}

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');

build.addSuppression(
  `Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`
);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(gulp);

const gulpSequence = require('gulp-sequence');

gulp.task(
  'all-in-one-go',
  gulpSequence('clean', 'build', 'bundle', 'package-solution')
);
