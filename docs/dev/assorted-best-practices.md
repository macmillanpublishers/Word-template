# Custom Ribbon tab / Mac toolbar

# Adding new styles
## Naming conventions
Each style name must end with a short style code in parentheses. Paragraph styles are first-letter caps, and character styles are all lowercase. Character styles all begin with the word "span".

## Formatting conventions
### macmillan.dotm
 Each paragraph style must have a colored left-side border added to its formatting, and each character style must have colored shading added to its formatting.

### macmillan_NoColor.dotm
Matches `macmillan.dotm` exactly, except that the left-side border and background shading have been removed.


# Non-template file dependencies
## Bookmaker check macro
Requires `Styles_Bookmaker.csv`

## Castoff
Currently in develop branch only. But requires castoff CSV files.

# Update version numbers
Copy `Utilities.dotm` into Word Startup dir. Be sure to add the path to your local git repo. Userform pulls "current" version from text file in repo. If you change it, it updates the version file, updates the internal version number in the local template file, then copies the local file back into the repo. Just add, commit, and push!


# Issues and labels


# Working with git
General git intro. Maybe repeat some stuff about why we export modules.

## Line endings

## Merge conflicts