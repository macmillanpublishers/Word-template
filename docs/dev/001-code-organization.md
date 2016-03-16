# General organization
VBA macros are run from code modules in Office template files, and our system requires multiple templates. Our general goal is to have the code for each macro in its own module, but as there are some exceptions and dependencies among modules and templates, we recommend that you use all of the templates together. 

We use Git for version control, [and the repository is hosted on Github](https://github.com/macmillanpublishers/Word-template). Git can't merge binary files, so we also export all code modules to the repo (but note that end-users only need the templates files). Binary template files are stored in subdirectories with the same name as the file, along with the code modules that live in that template (with the exception of modules shared by different templates--see below).

There is also a plain text file with the same name as each template file in its subdirectory, which stores that template's version number.


# Global templates and modules

## Shared modules
Modules used used in multiple template files are stored in the `SharedModules` subdirectory. The module name *must* start with `Shared`. ***NOTE:*** Managing changes to these files among the different templates can get tricky; see info about this later.

### SharedFileInstaller.bas
Contains code to check for template updates and download them if needed.

### SharedMacros.bas
This module holds procedures that are used by multiple individual macros. Procedures in macro-specific modules are defined as `Private`, meaning they can only be called from within the same module. In contrast, the procedures in this module are all defined as `Public` so they can be called from any module.


## MacmillanTemplateInstaller.docm
Installs the Macmillan template. See Installer docs for more info. Also requires `SharedFileInstaller.bas` and `SharedMacros.bas`. **Note: this won't work for in-house Mac users who do not have admin rights on their machines.**

### ThisDocument.cls
Contains a `Document_Open` procedure that runs (you guessed it) when the document is opened.



## GtUpdater.dotm
This is a Global Template that checks daily if the main macro template, `MacmillanGT.dotm`, is up to date and loaded as an Add-In. If not, it prompts the user and downloads the updated template. Also requires `SharedFileInstaller.bas` and `SharedMacros.bas`.

### ThisDocument.cls
Contains an `AutoExec` procedure that runs when `GtUpdater.dotm` is loaded as an Add-In. Calls `SharedFileInstaller.bas` and `SharedMacros.bas`.



## MacmillanGT.dotm
This is the main Global Template, and as such it stores the code for the macros people actually use. It also checks daily that `GtUpdater.dotm` and the style templates are up to date, and if not it prompts the user and downloads the updates.  Also requires `SharedFileInstaller.bas` and `SharedMacros.bas`.

### AttachTemplateMacro.bas
Attaches the Macmillan style templates (`macmillan.dotm`, `macmillan_NoColor.dotm`, `macmillan_CoverCopy.dotm`) to load their styles into the current document.  Requires `SharedMacros.bas`.

### CharacterStyles.bas
Converts direct formatting of italics, bold, small caps, etc. to Macmillan character styles and removes unstyled page breaks and blank paragraphs. Requires `ProgressBar.frm` and `SharedMacros.bas`.

### CleanupMacro.bas
Fixes common typographic errors. Requires `ProgresBar.frm` and `SharedMacros.bas`.

### EasterEggs.bas
Adds an ASCII triceratops to the end of the document.

### Endnotes.bas
Unlinks embedded endnotes and places them in a section at the end of the document.  Requires `ProgresBar.frm` and `SharedMacros.bas`.

### LOCtagsMacro.bas
Converts Macmillan-styled manuscripts to tagged text files for the Library of Congress CIP application following [these specifications](https://www.loc.gov/publish/cip/techinfo/formattingecip.html). Requires `ProgressBar.frm` and `SharedMacros.bas`.

### MacmillanCustomRibbonPC2007-2013.xml
Contains the code for the custom Ribbon tab, which can be added with the [Custom UI Editor for Microsoft Office.](http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2009/08/07/7293.aspx) Requires `RibbonControl.bas`.

### PrintStyles.bas
Adds paragraph style names to left margin for hardcopy printing. Requires `ProgressBar.frm` and `SharedMacros.bas`.

### ProgBarHelper.bas
Updates the progress bar (created via the `ProgressBar.frm` class) and waits until it's complete before allowing the rest of the main macro to run. `ProgressBar.frm` is a Userform that runs modeless (i.e., asynchronously), which can crash the macro if it hasn't finished updating before another call to further update happens. Requires `ProgressBar.frm` and `SharedMacros.bas`.

### ProgressBar.frm, ProgressBar.frx
Userform that displays a progress bar while other macros are running (PC only). Requires `ProgBarHelper.bas` and `SharedMacros.bas`.

### Reports.bas
Contains two reports that verify that the manuscript is styled following [Macmillan best practices](https://confluence.macmillan.com/display/PBL/Manuscript+Styling+Best+Practices). Requires `ProgressBar.frm` and `SharedMacros.bas`.

### RibbonControl.bas
Loads and controls the custom Ribbon tab. Requires `MacmillanCustomRibbonPC2007-2013.xml` via [Custom UI Editor](http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2009/08/07/7293.aspx).

### ThisDocument.cls
Contains `AutoExec` procedure to check daily if the `GtUpdater.dotm` template needs to be updated. Requires `SharedFileInstaller.bas` and `SharedMacros.bas` modules.

### VersionCheck.bas
Tells the user the current version number of `MacmillanGT.dotm` and `macmillan.dotm`, for troubleshooting purposes.

### ViewStyles.bas
Opens windows and page views useful for working with styles.




# Style templates
These templates store the Macmillan custom style sets. [Styles are listed here.](https://confluence.macmillan.com/display/PBL/Word+Template+Styles+List)

## macmillan.dotm
This is the primary style template, with color guides to make it easier to identify which styles are in use.

## macmillan_NoColor.dotm
This template contains all of the styles in `macmillan.dotm` with the same exact names and formatting, except the color guides have been removed. Used for printing bound manuscripts or not freaking out authors.

## MacmillanCoverCopy.dotm
This template contains all of the styles for jacket / cover copy. (We don't actually use these in our workflow right now, though.)



# Dev tools
## Utilities.dotm
A few macros just for working with VBA. To use, just copy this template file into your Word Startup directory and update the path to your local git repo in the private constant *strRepoPath* at top of the `Utilities.bas` module. As of now only tested on PC. 

### ThisDocument.cls
Contains two helpful macros: 

* `ExportAllModules` exports all code modules of open template to local git repo
* `ImportAllModules` imports all required modules from local git repo.

### VersionForm.frm, VersionForm.frx
Opens a Userform which displays the current version number of each template file (based on the version text file in the repo); can optionally update versions as well.