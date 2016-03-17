# Naming conventions
* Each style name must end with a short, *unique* style code in parentheses. 
* Paragraph style names are first-letter (of each word) caps
* Character style names are all lowercase. 
* Paragraph style names begin with the word indicating their section of the book.
* The only characters allowed are letters, numbers, spaces, and hyphens; # for space before/after/around styles, and parentheses surrounding the code. Do NOT use en- or em-dashes.



# Creating new styles
**Note:** Styles are stored in Word template files, and as such details of changes can't be tracked or merged with Git. So we need to take extra care that all branches have the same version of the style templates.

1. Make sure you have the most up-to-date `master` branch of the `Word-template` repo.
1. Checkout a new `hotfix-*` branch.
1. Copy *all* the style template files from the repo to the `MacmillanStyleTemplate` directory.
1. Launch Word, then use **File > Open** to open the file `macmillan.dotm` from the `MacmillanStyleTemplate` directory. (Do NOT click on the file to open it; this opens a new document based on the template, not the template itself.)
1. Open the **Styles Pane**.
1. Select an already-existing style in the body of the document to base the new style on.
1. Click the **New Style** button on the Styles Pane.
	1. On a Mac
	![Mac Styles Pane](images/mac-new-style.png)
	1. On a PC
	![PC Styles Pane](images/pc-new-style.png)
1. Create your style in the New Style window that opens.
![New Style Dialog Box](images/new-style-dialog.png)
	1. Add a new name, following the naming conventions for the Macmillan template.
	1. If you would like the style to inherit some of it's formatting from another style, select that style in the **Based on** drop-down menu.
	1. Set the formatting for the new style with the options in the window, or via the menus on the `Format` drop-down in the lower-left corner. If you are using a Based On style, you only need to set the formatting that you want to be *different* that the parent style.
	1. Include a left-margin colored border to paragraph styles or background shade for character styles.
	1. Deselect the "Add to Quick Styles list" button if it is selected.
	1. Click OK.
1. Now, in the body of the document, add the new style name in a new paragraph, *exactly* as you entered it in the New Style dialog window, and apply the new style to to that paragraph.
1. Add a note about what changes you made to the updates list at the top of the document.
1. Save and close the template.
1. Use the **File > Open** method to open `macmillan_NoColor.dotm` from the `MacmillanStyleTemplate` directory.
1. Repeat the steps above to add a style with the same *exact* name to `macmillan_NoColor.dotm`, but with no colored border or background shading.
1. Save and close the templates.
1. Open a new document and click the **Add Styles to Manuscript** button on the **Macmillan Tools** tab. Verify that the new style was imported, then close the document without saving.
1. Open the **Update Versions** utility macro. Enter the new version number in the userform that pops up and click **Update Versions**. 
1. Click **OK**. This will update the version number in the text file in the repo, update the version number in the template files' **Version** custom document property, and copy the updated template files to the repo.
1. Commit and push the changes.
1. Submit a pull request to the `Word-template` repo.
1. Checkout the `develop` branch.
1. Merge the `hotfix-*` branch into the `develop` branch, checkout the files from the `hotfix-*` branch if there are merge conflicts, and push the changes.
	
	```
	$ git checkout develop
	$ git merge --no-ff hotfix-*
	CONFLICTS ...
	$ git checkout --theirs macmillan/macmillan.dotm macmillan/macmillan_NoColor.dotm
	$ git commit -a -m "Fixing merge conflicts"
	$ git push origin develop
	```
1. Repeat the last step with the `releases` branch and any active features branches. 
1. Add the new style names to the [Word Template Styles List](https://confluence.macmillan.com/display/PBL/Word+Template+Styles+List) on Confluence.
1. If the style was not added in response to a request from Bookmaker, submit an issue to the [`bookmaker_assets` repo](https://github.com/macmillanpublishers/bookmaker_assets/issues) to add support for the style. The issue should be called: "New Style Request: Style Name Here", and in the body, make sure to add the style name exactly as it appears in the Word template, along with what kind of style it is (paragraph or character) and the reason it is being added (e.g., "request by Tor.com").