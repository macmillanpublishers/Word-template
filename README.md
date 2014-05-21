# Macmillan Word Styles and Template

This is a stub.

- The macmillan.dotm template file collects the Microsoft Word style set and macros for editorial production work.
- The MacmillanGT.dotm (GT = Global Template) will be a Word startup item, with a macro to attach the macmillan.dotm & load its styles.

## Dependencies

Word-macros

## Distribution End Points

Publishing Tools server share
Also available for client install on Macs via Casper 'Self Service' in Digital Workflow category.

## Deployment

**For Development**

On the server, cd into the directory and run:

```
$ git pull origin master
```

**Client installation: Mac**

*macmillan.dotm* goes here:  /Macintosh HD/Users/Shared/MacmillanStyleTemplate/

*MacmillanGT.dotm* goes here:  /Macintosh HD/Applications/Microsoft Office 2011/Office/Startup/Word/

Quick menu items ('Attach template' and 'View Macros' buttons) load along with the Global template at startup

**Client installation: PC**

*macmillan.dotm* goes here:  C:\ProgramData\MacmillanStyleTemplate\macmillan.dotm  

*MacmillanGT.dotm* goes either of these places:  

- C:\Program Files (x86)\Microsoft Office\Office14\STARTUP\  *(preferable, but requires admin permissions)*
- C:\Users\ *username* \AppData\Roaming\Microsoft\Word\STARTUP\   *alternate, account-based option*

Quick menu items ('Attach template' and 'View Macros' buttons) need to be manually created (for now):

- 1) Once MacmillanGT.dotm is in aforementioned STARTUP folder, relaunch Word to load it as a global template
- 2) goto File-> Options-> Quick Access Toolbar
- 3) Under 'Choose Commands from:' select 'Macros' and 'Add' desired one.
- 4) Under 'Choose Commands from:' select 'Popular' and 'Add' 'View Macros'.

## Stakeholders

Editorial Production, Design

## Usage

See confluence page.