'use strict';

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

const subtaskBuildChild1 = build.subTask(
  'subtask-buildChild1',
  function (gulp, buildOption, done) {
    console.log('sub-task-buildChild1 of build through console.log');
    this.log('sub-task-buildChild1 of build through this.log');

    this.logWarning('this is logWarning from sub-task-buildChild1');
    this.logError('this is logError from  sub-task-buildChild1');

    this.fileWarning(
      './dir1/subdir1/samplefile1.ts',
      15,
      20,
      1520,
      'warning code',
      'this is fileWarning from sub-task-buildChild1'
    );
    this.fileError(
      './dir1/subdir1/samplefile2.ts',
      25,
      20,
      2520,
      'error code',
      'this is fileError from sub-task-buildChild1'
    );

    done();
  }
);

build.task('subtask-buildChild1', subtaskBuildChild1);

const subtaskBuildChild2 = build.subTask(
  'subtask-buildChild1',
  function (gulp, buildOption, done) {
    this.log('sub-task-buildChild2 of build through this.log');
    done();
  }
);

build.task('subtask-buildChild2', subtaskBuildChild2);

build.initialize(gulp);

if (gulp.tasks['build']) {
  gulp.tasks['build'].dep.push('sub-task-buildChild1', 'sub-task-buildChild2');
}
