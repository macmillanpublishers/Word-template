When a macro-enabled *document* (.docm) is opened, the user can run any of the macros it contains. When a macro-enabled *template* (.dotm) is loaded as an Add-In, it is considered a Global Template and the user can run its macros on any document. 


To view the modules (and code) in a template, type `Alt` + `F11` to open the Visual Basic Editor (VBE). Modules are shown in the Project window of the VBE.


# Opening/closing template files
# Always work on the local file
Because local version will be loaded automatically and can be confusing which you're interacting with. Use Export/Import Utilities to copy to repo.

## Use File > Open only
Double click opens a new doc based on that template.

## Saving template files
Don't save on a Mac--specifically `MacmillanGT.dotm` because it removed custom XML, but in general just to be safe. Instructions for how to export modules, commit and push, then pull to PC, Import all modules, then commit and push again.

# Import / Export code  modules
Copy `Utilities.dotm` to your Startup folder. 

## Macros
### ExportAllModules
What it does in detail and issues it has. Probably will remove userform export (definitely for binaries)
#### Export modules manually

### ImportAllModules
Ditto
#### Import modules manually


# Userforms
UserForms appear in the template as a single module, but when exported two files are created: `ProgressBar.frm` is a class module that contains properties and methods for this class, while `ProgressBar.frx` is a binary file containing the design of the Userform itself. You only need to select the .frm file to import, but both must be present in the same directory.


# AutoExec and Document_Open procedures
Must be in `ThisDocument.cls` module, which incidentally can't be removed, because it's the object for the template file itself.