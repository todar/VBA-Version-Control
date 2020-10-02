# VBA Version Control

<a href="https://www.buymeacoffee.com/todar" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/default-orange.png" alt="Buy Me A Coffee" style="height: 51px !important;width: 217px !important;" ></a>

This module is created to easily export and import VBA code to `./src` directory. This gives the ability to then use Git to version control your VBA 🎉!

This is needed as Git can't read Excel files directly, but can read the source files that are exported.

> Added bonus, there is a `Bash` function to run commands directly within the Immediate window. Example `Project.bash "git add . && git commit -m ""Yay, from VBA"""

## Goal of the project

- Create an easy way to integrage git version control with VBA.
- Create continuous integration for VBA Projects.

## Progress

- [x] Write code to export VBComponents to source directory.
- [x] Write code to import VBComponents from source directory.
- [ ] Create a testing workflow for continuous integration/deployment.

## Getting Started

First make sure you have [git](https://git-scm.com/) installed on your system.

1. import `Project.bas` into your project.
1. Set a reference to `Microsoft Visual Basic For Applications Extensibility 5.3` and `Microsoft Scripting Runtime`
1. In the immediate widow type `Project.InitializeProject`. This will create a gitignore file and run `git init`.
1. Run `Project.ExportComponentsToSourceFolder` manually, or even better add it to the `ThisWorkbook.Workbook_AfterSave` event, so that it runs after every save.

That's it! Version control can now be managed easily with VBA.
