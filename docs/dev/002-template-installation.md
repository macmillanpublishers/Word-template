# Overview
End users can install the template themselves following the instructions [on this page](https://confluence.macmillan.com/display/PBL/Install+the+Macmillan+Template). This document just explains what the installer file does for troubleshooting or manual installation purposes (or just because you're curious).

# Installed file locations
There are a number of template files that need to be saved in specific directories for the macros to work correctly.

## Startup
Word automatically loads any templates saved in the Startup directory  as Add-Ins when it launches. 

### Files to be installed here
* **GtUpdater.dotm:** Has an `AutoExec` procedure that loads `MacmillanGT.dotm` as an Add-In and also checks daily that `MacmillanGT.dotm` is up to date, then prompts the user and downloads the updated files if needed.

### Location
Note that because the Startup location can be changed by the user, the aths listed below are just the defaults. The installer macro actually uses the `Application.StartupPath` property to locate it.

#### PC
* `C:\Users\<username>\AppData\Roaming\Microsoft\Word\STARTUP` (where `<username>` is replaced by the actual username)

You can change the current Startup directory in Word 2010 and 2013 via **File > Options > Advanced > General > File Locations... > Startup**.

#### Mac
* `Macintosh HD:Applications:Microsoft Office 2011:Office:Startup:Word`

You can change the current Startup directory in Word 2011 via **Preferences > File locations > Startup**.

Note that the Startup directory is read-only for Mac users without admin access, which is why Mac users in-house install the template via use Self Service.


## MacmillanStyleTemplate
All other templates must be stored in the `MacmillanStyleTemplates` directory. 

Note that the directory name must be spelled exactly as listed (camel case, no spaces) for the macros to work correctly.


### Location
#### PC
The MacmillanStyleTemplate directory is located in C:\ProgramData\MacmillanStyleTemplate (the root being of course the user's local drive). If it does not already exist the installer file will create it. Note that ProgramData is a hidden directory. 

#### Mac
The MacmillanStyleTemplate directory is located in the user's Documents directory (Macintosh HD:Users:<username>:Documents). If it does not already exist the installer file (or Self Service) will create it. 


### Files to be installed here
* **MacmillanGT.dotm:**the primary template containing the Macmillan custom macros and custom ribbon tab.
* **macmillan.dotm:** the primary style set, containing over 400 custom manuscript styles.
* **macmillan_NoColor.dotm:** contains exactly the same styles as `macmillan.dotm` but with colored borders and shading removed.
* **macmillan_CoverCopy.dotm:** contains styles for use in jacket / cover copy documents


# Other files you may see
These files may appear in the same directories as the template files, but do not need to be available for the template to work.

## Version files
Each template has a text file of the same name, which contains the current version number of that template. When the templates check for updates each day they save that file in `MacmillanStyleTemplates'.

## Log files
A subdirectory named `log` is created by the installer macro in the `MacmillanStyleTemplate` directory. It contains logs of downloads for each template file.

 The Castoff macro also saves CSV files of design info here, and the Bookmaker check macro saves a CSV of current 


# Installer file
The download for users is linked directly to the installer file hosted on the Macmillan Publishers GitHub account in the master branch of the word-template repo. Any changes pushed to master are therefore immediately available to users for the installer file, but not for the template files. The template files by contrast are hosted as attachments on this page. There is also a staging page to use for testing.

The installer file launches the installation macro immediately at file open, though with default security settings most users will be prompted with a warning and need to click a button to enable macros before it will run.

If you need to open the file without running the macro (to edit code, for example), hold down Shift while you open it.


# Update checker


# Git Confluence connector
What it is and how it works.
Something about branches.

## Scheduled jobs

