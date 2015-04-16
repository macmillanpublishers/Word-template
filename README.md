# Macmillan Word Styles and Templates

The Macmillan template files collect the Microsoft Word style sets, macros, and macro launch buttons for editorial production work. All macros are launched from custom Quick Access Toolbar buttons (Windows) and custom toolbars (Mac) saved in each template.

## MacmillanGT.dotm

MacmillanGT.dotm (GT = Global Template) is a Word startup item. It contains the following macros:

### AttachTemplateMacros.bas

Attaches the Macmillan style templates (macmillan.dotm, macmillan_NoColor.dotm, MacmillanCoverCopy.dotm) & loads their styles and macros into the current document. [More info.](http://confluence.macmillan.com/display/PBL/Attach+the+Macmillan+Template)

### VersionCheck.bas

Tells the user the current version number of MacmillanGT.dotm and macmillan.dotm. [More info--requires login.](http://confluence.macmillan.com/pages/viewpage.action?pageId=32112870)

## macmillan.dotm

The primary Macmillan style template. It contains all of the [custom styles](http://confluence.macmillan.com/display/PBL/Word+Template+Styles+List) for use in Macmillan manuscripts, and the following macros:

### CharacterStyles.bas

Converts direct formatting of italics, bold, small caps, etc. to Macmillan character styles and removes unstyled page breaks and blank paragraphs. [More info.](http://confluence.macmillan.com/display/PBL/Macmillan+Character+Styles+Macro)

### CleanupMacro.bas

Fixes common typographics errors. [More info.](http://confluence.macmillan.com/display/PBL/Macmillan+Manuscript+Cleanup+Macro)

### Reports.bas

Contains two reports that verify that the manuscript is styled following [Macmillan best practices.](http://confluence.macmillan.com/display/PBL/Manuscript+Styling+Best+Practices)

#### Style Report

Produces a report listing all of the Macmillan paragraph styles used and the page number and paragraph number of any non-Macmillan styles used. [More info.](http://confluence.macmillan.com/display/PBL/Macmillan+Style+Report+Macro)

#### Bookmaker Requirements

Produces a report listing errors in the manuscript that need to be resolved before using the Macmillan Bookmaker tool. [More info--requires login.](http://confluence.macmillan.com/display/PE/Bookmaker+Requirements+Macro)

### ViewStyles.bas

Opens windows and page view useful for working with styles. [More info.](http://confluence.macmillan.com/display/PBL/View+Styles+with+a+Macro)

## macmillan_NoColor.dotm

Contains all of the same styles and macros as macmillan.dotm but without colored shading and borders.

## MacmillanCoverCopy.dotm

Contains custom styles for jacket and cover copy. Also includes Cleanup, Style Report, and View Styles macros.

## LOC_Tags.dotm

Contains LOCtagsMacro.bas, a macro that converts Macmillan-styled manuscripts for the Library of Congress CIP application.

## SwoonReadsTemplate.dotm

Contains the limited Macmillan style set for Swoon Reads.

## torCastoffTemplate.dotm

Contains torCastoffMacro.bas, a macro that estimates the print page count of manuscripts for Tor.com.

# Dependencies

Macros are stored in modules in the Visual Basic Editor in each template, and also exported as .bas files and stored in the same directory as their source template. UserForms are exported as .frx files.

To export macros and userforms from the template files, ctrl+click/right-click on the module and selecting "Export file...". 

To import macros and userforms into the template files, ctrl+click/right-click on the module in the template and select "Remove [name of module]", then ctrl+click on "TemplateProject(NameOfTemplate)", select "Import file...", and select the .bas file of the module to import.

Any style or macro changes made to macmillan.dotm must be made to macmillan_NoColor.dotm.

Templates must be saved on Windows; saving on a Mac causes the Windows macro launch buttons to drop out.

# Distribution End Points

## MacmillanGT.dotm

Direct link to raw file on GitHub on [Macmillan Confluence site.](http://confluence.macmillan.com/display/PBL/Install+the+Macmillan+Template#InstalltheMacmillanTemplate-OnaPC(orMac-external))

Also available for client install on Macs via Casper 'Self Service' in Digital Workflow category.

## macmillan.dotm, macmillan_NoColor.dotm, MacmillanCoverCopy.dotm

Attachments to [public Macmillan Confluence page;](http://confluence.macmillan.com/display/PBL/Test).

## LOC_Tags.dotm

Users download from direct link to raw file on GitHub on [Macmillan Confluence site--requires login.](http://confluence.macmillan.com/display/PE/CIP+Tagging+Macro)

## torCastoffTemplate.dotm

Users download from direct link to raw file on GitHub on [Macmillan Confluence site--requires login.](http://confluence.macmillan.com/display/EDIT/Tor.com+Castoff+Macro)

# Deployment

[Word styles template update process--requires login.](http://confluence.macmillan.com/display/~erica.warren/Word+Styles+template+update+process)

# Client installation

## MacmillanGT.dotm

Users can download from direct link to raw file on GitHub on [Macmillan Confluence site.](http://confluence.macmillan.com/display/PBL/Install+the+Macmillan+Template#InstalltheMacmillanTemplate-OnaPC(orMac-external))

[Manual installation instructions.](http://confluence.macmillan.com/display/PBL/Install+the+Macmillan+Template)

Also available for client install on Macs via Casper 'Self Service' in Digital Workflow category.

## macmillan.dotm, macmillan_NoColor.dotm, MacmillanCoverCopy.dotm

Attachments to [public Macmillan Confluence page;](http://confluence.macmillan.com/display/PBL/Test) downloaded to users' machines via the AttachTemplateMacros.bas macros.

## LOC_Tags.dotm

[Manual installation instructions--requires login](http://confluence.macmillan.com/display/PE/CIP+Tagging+Macro#CIPTaggingMacro-InstallingtheMacro)

## torCastoffTemplate.dotm

[Manual installation instructions--requires login](http://confluence.macmillan.com/display/EDIT/Tor.com+Castoff+Macro)

# Stakeholders

Production Editorial, Design

# Usage

[More info available on the Macmillan Confluence site.](http://confluence.macmillan.com/display/PBL/About+the+Macmillan+Word+Template)