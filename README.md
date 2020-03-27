# VBA

VBA Version control using Git and npm.

> This is only in the concept and discovery phase!

## Goal of the project

- Create an easy way to integrage git version control with VBA.
- Use npm to build new VBA projects from source code.
- Create continuous integration for VBA Projects.

## Progress

- [x] Write code to export VBComponents to source directory.
- [ ] Write code to import VBComponents from source directory.

## Notes

- With imports, should I clear out all other code so that the import is clean? Possible danger of losing code that has not been saved yet. Possible danger of leaving dirty code... Not sure yet...

## How to use

While under development, run `npm run link` to make the project global. Then use `vba` in the command line.
