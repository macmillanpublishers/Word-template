Word-macros
==============
These vba macros are part of the IT Word Styles Project.  This repo is for macro vba *code* only; the related 'compiled' .dotm files are being maintained in the Macmillan Word-template repo.

Dependencies
--------------
Macmillan Word Styles and Template

Distribution End Points
--------------
Publishing Tools server share

Deployment
--------------
**Cleanup Macro**
- In MS Word, File->open the macmillan.dotm (available from Macmillan Word-template repo).
- Go to Tools->Macro->Visual Basic Editor. 
- In the Project Explorer you should see a project called: "TemplateProject(macmillan.dotm)"  The only Object listed under the project should be ThisDocument.  Double click on that to open the Code window for ThisDocument
- Select all/any existing code and delete.  Copy and Paste all code from latest version on this repo into the code window.
- With the macmillan.dotm Code window still active, goto File->Save to save the changes you made.
- Goto 'Word' menu, select 'Close and return to Microsoft Word'.  
- File->Save the macmillan.dotm document and close.

**Attach Template Macro**
- In MS Word, File->open the MacmillanGT.dotm (available from Macmillan Word-template repo).
- Go to Tools->Macro->Visual Basic Editor. 
- In the Project Explorer you should see a project called: "Macmillan(MacmillanGT.dotm)"  The only Object listed under the project should be ThisDocument.  Double click on that to open the Code window for ThisDocument
- Select all/any existing code and delete.  Copy and Paste all code from latest version on this repo into the code window.
- With the MacmillanGT.dotm Code window still active, goto File->Save to save the changes you made.
- Goto 'Word' menu, select 'Close and return to Microsoft Word'.  
- File->Save the MacmillanGT.dotm document and close.
- **Quick Menu Items**: These are platform independent, so must be created using both PC-Word and Mac-Word.
  - 1) For Macs: With the MacmillanGT.dotm still open, goto View->Toolbars->Customize Toolbars.. 
  - 2) Make sure the 'Save In' selected is MacmillanGT.dotm.  Select 'Commands' tab, Category 'Macros'
  - 3) Drag the Macro to the Quick menu toolbar.  If desired also drag "Macros..." from 'Tools' category to Quick menu bar as well.
  - 4) Save the document
- For Windows:
  - 1) With the MacmillanGT.dotm still open, goto File-> Options-> Quick Access Toolbar
  - 2) Under 'Choose Commands from:' select 'Macros' and 'Add' desired one.
  - 3) Under 'Choose Commands from:' select 'Popular' and 'Add' 'View Macros'.
  - 4) Save the document

Stakeholders
--------------
Editorial Production, Design

Usage
--------------
Macmillan Word Styles and Template repo, related confluence page
