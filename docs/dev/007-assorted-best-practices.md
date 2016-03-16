# Custom Ribbon tab / Mac toolbar



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
Git cannot merge changes to binary files, so if you have merged different branches you will get a merge conflict. Luckily you exported your code modules, which *can* merge two different branches, so just follow the steps below to update the template files.

From the command line, after you get a merge conflict:

```
$ git checkout develop # or whatever branch you are merging changes into
$ git merge feature # "feature" being whatever feature branch you are merging in; you may get a merge conflict here
```
Resolve any conflicts in the macro files (.bas, .cls, or .frm) and commit the changes.

```
git checkout --ours path/file.dotm (keeps the develop version of the file; if you want to keep the feature branch version, use --theirs)
git add path/file.dotm
```

Open the template file and replace the modules that were updated in feature with the modules currently in the repo (that have been merged) – the ImportAllModules macro in the Utilities.dotm template is useful for this part.
Save the template.
Add the modified file and commit the changes.