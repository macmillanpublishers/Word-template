# The Macmillan Word Template
The Macmillan Word template contains a variety of VBA macros for use in working with manuscripts. When installed on a PC, the user gets a custom "Macmillan Tools" tab on the ribbon with buttons to launch the macros; when installed on a Mac, the user gets a custom floating toolbar.

# Supported versions of Word
* **PC:** Tested and supported for 2007 - 2013. 
* **Mac:** Tested and supported for 2011 only.

We haven't tested anything on Office 2016 yet, but we're working on it!

# Installation
Download the Word-template.dotm file and double click the icon to open a new file. This will launch the installation macro.

# Dependencies
The Macmillan template uses the excellent [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) and [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary) projects.

# Config files
`global_config.json` contains data the template needs to manage file downloads and updates. The `files` key contains info for each file the template needs.

The values for the `global_config.json` file itself are also stored in the template's `CustomDocumentProperties`.

Files will be downloaded from the commit tagged with the value in the `latest_release` key, unless there is a `branch` key present in which case the file will be downloaded from the HEAD of that branch. If neither is present, the latest file on the `master` branch used.

Users are prompted to select their region; data in the region config files takes precedence over data in the global config. This is where you should specify which release that region should be on.

The user's region is stored in a local config file. If the local config file also contains a `files` key, those values take precedence over the global and region config files. 

# Releases
The current version number should be stored in the template's `CustomDocumentProperties` with the key `Version`. The version number itself should be prefixed with a `v` followed immediately by the version number: `v2.9.5`

When a new version is ready for release, the `latest_version` should be updated in `global_config.json` (and any relevant region config files). Once merged with the `master` branch, open a new Release in GitHub using the same version number string as the tag. You do not need to attach any additional binary files to the Release.

# Using the Macmillan Tools macros
Documentation for *using* the Macmillan template (and styles more generally) is available [here](https://confluence.macmillan.com/display/PBL/Manuscript+Styling+with+MS+Word).
