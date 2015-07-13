# Macmillan Word Styles and Templates

The Macmillan template files collect the Microsoft Word style sets, macros, and custom Ribbon tab or toolbar for editorial production work. Except for the Installer macro (which launches on file open), all macros are launched from a custom Ribbon tab (Windows) or a custom toolbar (Mac) saved in the MacmillanGT.dotm template. There are manu dependencies among the different modules so it is suggested that all be installed in the appropriate templates.


## Installers/

### MacmillanTemplateInstallerPC.docm

Installs MacmillanGT.dotm and GtUpdater.dotm in the Word Startup directory, creates MacmillanStyleTemplate directory (if needed), and installs macmillan.dotm, macmillan_NoColor.dotm, and MacmillanCoverCopy.dotm there. Contains the following macro, which launches when the user opens the file:

### ThisDocument.cls

Should be loaded into the ThisDocument module of MacmillanTemplateInstallerPC.docm as a sub titled Document_Open(). Requires FileInstaller module. 


## macmillan/

### macmillan.dotm

The primary Macmillan style template. It contains all of the [custom styles](https://confluence.macmillan.com/display/PBL/Word+Template+Styles+List) for use in Macmillan manuscripts.

### macmillan_NoColor.dotm

Contains all of the same styles and macros as macmillan.dotm but without colored shading and borders.

### MacmillanCoverCopy.dotm

Contains custom styles for jacket and cover copy.



## MacmillanGT/

### MacmillanGT.dotm

MacmillanGT.dotm (GT = Global Template) is a Word startup item. It contains the icons for the custom Ribbon tab, added with the [Custom UI Editor for Microsoft Office](http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2009/08/07/7293.aspx). It also contains the following items:

### MacmillanCustomRibbonPC2007-2013.xml

Contains the code for the custom Ribbon tab, which should be added with the Custom UI Editor. Requires the RibbonControl module to be imported into the same template.

### AttachTemplateMacros.bas

Attaches the Macmillan style templates (macmillan.dotm, macmillan_NoColor.dotm, MacmillanCoverCopy.dotm) to load their styles into the current document. [More info.](https://confluence.macmillan.com/display/PBL/Attach+the+Macmillan+Template)

### CastoffForm.bas, CastoffForm.frm, CastoffForm.frx

Creates a castoff for either SMP or Tor.com titles. Downloads design character counts from [Confluence page.](https://confluence.macmillan.com/display/PBL/Test) Also contains CastoffForm.frm, a userform to allow user to select trim size and relative type design (loose, average, tight).

### CharacterStyles.bas

Converts direct formatting of italics, bold, small caps, etc. to Macmillan character styles and removes unstyled page breaks and blank paragraphs. [More info.](https://confluence.macmillan.com/display/PBL/Macmillan+Character+Styles+Macro) Requires ProgressBar.frm.

### CleanupMacro.bas

Fixes common typographics errors. [More info.](https://confluence.macmillan.com/display/PBL/Macmillan+Manuscript+Cleanup+Macro). Requires ProgresBar.frm.

### LOCtagsMacro.bas

Converts Macmillan-styled manuscripts to tagged text files for the Library of Congress CIP application. Requires ProgressBar.frm.

### ProgressBar.frm, ProgressBar.frx

Userform that displays a progress bar as other macros are running.

### Reports.bas

Contains two reports that verify that the manuscript is styled following [Macmillan best practices.](https://confluence.macmillan.com/display/PBL/Manuscript+Styling+Best+Practices). Requires ProgressBar.frm.

### RibbonControl.bas

Contains macros to load and control the custom Ribbon tab.

### ThisDocument.cls

Contains an AutoExec macro to check daily if the GtUpdater.dotm template needs to be updated. Must be saved in the ThisDocument object module as a sub called AutoExec(). Requires FileInstaller module.

### VersionCheck.bas

Tells the user the current version number of MacmillanGT.dotm and macmillan.dotm. [More info--requires login.](https://confluence.macmillan.com/pages/viewpage.action?pageId=32112870)

### ViewStyles.bas

Opens windows and page view useful for working with styles. [More info.](http://confluence.macmillan.com/display/PBL/View+Styles+with+a+Macro)


# Dependencies

Macros are stored in modules in the Visual Basic Editor in each template, and also exported as .bas files and stored in the same directory as their source template. UserForms are exported as .frx and .frm files.

To export macros and userforms from the template files, ctrl+click/right-click on the module and select "Export file...". 

To import macros and userforms into the template files, ctrl+click/right-click on the module in the template and select "Remove [name of module]", then ctrl+click on "TemplateProject(NameOfTemplate)", select "Import file...", and select the .bas file of the module to import.

Any style or macro changes made to macmillan.dotm must be made to macmillan_NoColor.dotm.

Templates must be saved on Windows; saving on a Mac causes the Windows custom Ribbon code and icons to drop out.


# Distribution End Points

## MacmillanTemplateInstaller.docm

Direct link to raw file on GitHub on [Macmillan Confluence site](https://confluence.macmillan.com/display/PBL/Install+the+Macmillan+Template). For PC and external Mac installation only (requires admin rights on a Mac). 

## MacmillanGT.dotm

Saved as an attachment on top-level page of [Public Macmillan Confluence site](https://confluence.macmillan.com/display/PBL/Test) for download by the Installer or update via GtUpdater.dotm.

Also available for client install on internal Macs via Casper 'Self Service' in Digital Workflow category.

## GtUpdater.dotm

Saved as an attachment on top-level page of [Public Macmillan Confluence site](https://confluence.macmillan.com/display/PBL/Test) for download by the Installer or update via MacmillanGT.dotm.

Also available for client install on internal Macs via Casper 'Self Service' in Digital Workflow category.

## macmillan.dotm, macmillan_NoColor.dotm, MacmillanCoverCopy.dotm

Attachments to [public Macmillan Confluence page](http://confluence.macmillan.com/display/PBL/Test), for download via the Installer or update via MacmillanGT.dotm.


# Deployment

[Word styles template update process--requires login.](http://confluence.macmillan.com/display/~erica.warren/Word+Styles+template+update+process)

# Client installation

## PC and external Mac: MacmillanTemplateInstaller.docm

Users can download from direct link to raw file on GitHub on [Macmillan Confluence site.](http://confluence.macmillan.com/display/PBL/Install+the+Macmillan+Template#InstalltheMacmillanTemplate-OnaPC(orMac-external))

## Mac: Self Service

Available for client install on Macs via Casper 'Self Service' in Digital Workflow category. Installs all other required templates.

## macmillan.dotm, macmillan_NoColor.dotm, MacmillanCoverCopy.dotm

Attachments to [public Macmillan Confluence page;](http://confluence.macmillan.com/display/PBL/Test) downloaded to users' machines via the MacmillanGT.dotm AutoExec macro.


# Stakeholders

Production Editorial, Design

# Usage

[More info available on the Macmillan Confluence site.](http://confluence.macmillan.com/display/PBL/About+the+Macmillan+Word+Template)