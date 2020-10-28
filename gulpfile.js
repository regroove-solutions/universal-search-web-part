"use strict";

const build = require("@microsoft/sp-build-web");

build.addSuppression(
  `Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`
);

//https://n8d.at/how-to-version-new-sharepoint-framework-projects/
//synchronize manifest version with package version
let syncVersionsTask = build.subTask("version-sync", function (
  gulp,
  buildOptions,
  done
) {
  const fs = require("fs");
  var pkgConfig = require("./package.json");
  var pkgSolution = require("./config/package-solution.json");
  var teamsManifest = require("./teams/manifest.json");

  var newVersionNumber = pkgConfig.version.split("-")[0];

  if (teamsManifest.version !== newVersionNumber) {
    teamsManifest.version = newVersionNumber;

    this.log("Sync build version (Teams):\t" + teamsManifest.version);

    fs.writeFile(
      "./teams/manifest.json",
      JSON.stringify(teamsManifest, null, 2),
      function (err, result) {
        if (err) this.log("error", err);
      }
    );
  }

  newVersionNumber += ".0";
  if (pkgSolution.solution.version !== newVersionNumber) {
    pkgSolution.solution.version = newVersionNumber;

    this.log("Sync build version (package):\t" + pkgSolution.solution.version);

    fs.writeFile(
      "./config/package-solution.json",
      JSON.stringify(pkgSolution, null, 2),
      function (err, result) {
        if (err) this.log("error", err);
      }
    );
  }
  done();
});
build.rig.addPreBuildTask(build.task("version-sync", syncVersionsTask));

build.initialize(require("gulp"));
