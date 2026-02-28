// 'use strict';
// require("dotenv").config();
// const build = require('@microsoft/sp-build-web');

// build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// var getTasks = build.rig.getTasks;
// build.rig.getTasks = function () {
//   var result = getTasks.call(build.rig);
//   result.set('serve', result.get('serve-deprecated'));
//   return result;
// };

// build.initialize(require('gulp'));


'use strict';

require("dotenv").config();

const build = require('@microsoft/sp-build-web');
const webpack = require('webpack');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// Inject environment variables into SPFx bundle
build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {

    generatedConfiguration.plugins.push(
      new webpack.DefinePlugin({
        'process.env.AZURE_OPENAI_API_KEY': JSON.stringify(process.env.AZURE_OPENAI_API_KEY)
      })
    );

    return generatedConfiguration;
  }
});

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);
  result.set('serve', result.get('serve-deprecated'));
  return result;
};

build.initialize(require('gulp'));