#!/usr/bin/env node

/**
 * This is the main command line interface that will be able to
 * run all the commands such and import, export, and build.
 *
 * Please note, this is only concept so far, not ready for
 * use in production.
 *
 * @see https://github.com/tj/commander.js
 */

const path = require("path");
const inquirer = require("inquirer");
const { program } = require("commander");
const { spawn } = require("child_process");
// const shell = require("shelljs");
// const chalk = require("chalk");

// Sample prompt style for running commands
// Not sure if this will be used, but might be helpful if
// there are several steps to creating a new project...
async function prompt() {
  const questions = [
    {
      type: "list",
      name: "action",
      message: "Please pick an action:",
      choices: ["Import", "Export", "Build"],
      filter: function(val) {
        return val.toLowerCase();
      },
    },
  ];
  const answers = await inquirer.prompt(questions);
  if (answers.action === "export") {
    exportComponents();
  } else {
    console.log("Not developed yet!");
  }

  console.log(JSON.stringify(answers, null, "  "));
}

// Call to export.vbs script to run the exportComponents. This is exported
// to the src folder in the root directory.
function exportComponents() {
  const wscriptFilePath = path.join(
    "C:",
    "Windows",
    "SysWOW64",
    "wscript.exe"
  );
  const scriptPath = path.join(__dirname, "export.vbs");
  const appPath = path.join(__dirname, "../", "App.xlsm");
  const bat = spawn(wscriptFilePath, [scriptPath, appPath]);
}

// Basic details for the CLI.
program
  .version(
    // @ts-ignore
    require("../package.json").version,
    "-v, --version",
    "output the current version"
  )
  .name("vba")
  .description(
    "A tool for managing Excel VBA projects with npm and Git."
  );

// Call to the sample prompt above... Again not sure if this will be used.
program
  .command("prompt")
  .description("A sample prompt")
  .action(prompt);

// Exports components from the main Excel App into the src directory.
program
  .command("export")
  .description("Export all components from App.xlsm to src")
  .action(exportComponents);

// Let the commander program parse any arguments from the user.
program.parse(process.argv);

// Display the default help options if no argument was passed in.
if (!process.argv.slice(2).length) {
  program.help();
}
