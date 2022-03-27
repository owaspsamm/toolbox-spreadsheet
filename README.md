# SAMM Toolbox Spreadsheet Generator

This GitHub Action generates the SAMM toolbox spreadsheet based on the core model. 

**If you're looking for the spreadsheet itself, you'll find the newest version at https://github.com/owaspsamm/core/releases/latest.**

## How This GitHub Action Works
The action is called by the [core model build process](https://github.com/owaspsamm/core/blob/develop/.github/workflows/release.yml). It expects the updated model files in the `/github/workspace/model` directory of the runner. 

It loads these files, opens the "dummy" spreadsheet located in the `resources` folder in this repository and replaces the relant content. In the end, it stores the updated spreadsheet under `/github/workspace/SAMM_spreadsheet.xlsx` on the runner where it's taken by the parent build process to be published among the release deliverables.
