# Custom Ribbon tab / Mac toolbar
The macros in the template are launched via a custom Ribbon tab on PCs and a custom floating toolbar on Mac. The custom Ribbon tab is created by converting the template to a .zip file, unzippping it, and then adding a `customUI\customUI.xml` file. That's a hassle; luckily you can just use the [Custom UI Editor for Microsoft Office](http://openxmldeveloper.org/blog/b/openxmldeveloper/archive/2009/08/07/7293.aspx), which also makes it easy to add images for buttons.

The XML code is also saved directly in the repository as `MacmillanGT/MacmillanCustomRibbonPC2007-2013.xml`.

Here's info on how to [structure the XML](https://msdn.microsoft.com/en-us/library/aa942866.aspx?cs-save-lang=1&cs-lang=vb#code-snippet-3), but note that for the Macmillan template to work, the full macro name (module and procedure) MUST be stored as the `<button ID>` attribute.

A PC creates the tab from the XML directly, but to run the macros you need to include callbacks via the `RibbonControl.bas` module.

Word for Mac 2011 does not support custom XML, but it does support custom toolbars. To make this easier to maintain, `MacmillanGT.ThisDocument` contains a procedure that creates the toolbar by reading the same XML files the PC uses.

# Compiler constants
To make different code for different operating systems or application versions, use a `#` before the `#If... Then... #Else` statement. The code will only be compiled if it passes the test (so you won't get compile errors on code that won't run on that system anyway). Further info and a list of available constants to test are listed [here](https://msdn.microsoft.com/en-us/library/office/gg264614.aspx). Some common ones are:

* `Mac`: Any Mac operating system
* `Win16`: 16-bit Windows
* `Win32`: 32-bit Windows
* `Win64`: 64-bit Windows

Example:
```
#If Mac Then
  ' Mac specific code
#Else
  ' PC specific code
#End If
```

# Version numbers
If you need to test which version of Office is running, you can use the `Application.Version` property. The correct versions are:

| Office year (OS) | Version number |
| -------------- | ---------------- |
| 2016 (PC and Mac) | 16 |
| 2013 (PC) | 15 |
| 2011 (Mac) | 14 |
| 2010 (PC) | 14 |
| 2008 (Mac)* | 12 |
| 2007 (PC)** | 12 |
| 2003 (PC) | 11 |

* Office for Mac 2008 does not support VBA at all.
* Office 2007 is the first version to use the Ribbon interface.
 
# Office for Mac 2011
Office for Mac 2011 does not have the same object model as the PC versions do--much is the same, but the differences aren't documented anywhere consistently. Often this is overcome using the `MacScript()` function and compiler constants to run shell commands.

## 33-character file name limit
VBA in Office for Mac 2011 can't handle file paths longer than 32 characters, including the extension. The newer Office 2016 has overcome this, so test for the application version before calling any file functions such as `Kill()`, `Dir()`, `MkDir()`, and such. In fact, we're already created many custom functions for these in the `SharedMacros_` module.

# Update version numbers
The auto-update macro is triggered by a new version-number text file being greater than the `Version` custom document property in the template. To make it easier to maintain the various files, the `Utilities.dotm` template contains a macro specifically for managing version numbers. See [here](using+the+vb+editor) for info about installing it.

To check or change version numbers, click the red box button in the Quick Access Toolbar. A Userform will open displaying the current version number for each template. If you would like to change the number, enter the new version number in the text box provided and click **Update Versions**. This will update the `Version` custom document property in the template file, copy the template file to the repository, and update the version number text file in the repository. Now just commit the changes!



# Working with Git and Github
We use Git for version control, and use [Github](https://github.com/macmillanpublishers/Word-template) as our remote repository. This assumes a basic familiarity with Git, but there are some specific issues with working in VBA (for example, code runs from within a binary file) that we address here.

## Line endings
The Word template is designed to work cross-platform; however, working with line endings can be tricky. The code modules running in the template files need to use Windows-style line endings (CRLF) regardless of what platform the code is running on. Including any Unix-style (LF) line endings will prevent any code from running.

To sort this out, the `Word-template` repo contains a `.gitattributes` file which defines end-of-line characters for .bas, .cls, .frm, and .xml files to be CRLF in the working directory (Git converts all line endings to LR in the index). This should maintain the correct line endings regardless of platform.

If you want to be extra sure, you can update your Git configuration settings based on OS. Run the commands below to have PC check out code into the working directly with CRLF, and to have Mac check out code with LF.

* **PC:** - `git config --global core.autocrlf true`
* **Mac:** - `git config --global core.autocrlf input`




## Issues and labels
We use the [Issues](https://guides.github.com/features/issues/) feature on Github to manage bugs, new features, and the like. [You can see them here](https://github.com/macmillanpublishers/Word-template/issues)! Because the issues vary widely with regard to the amount of work they require and how important they are, we have a specific list of labels we use to make it easier to prioritize tasks. They are separated into three groups, and each issue should get one label from each group.

### Types of labels
#### Priority
There are four options for priority, which indicates how urgent the change is.

* **priority:critical** - Highest priority; must be implemented ASAP, even if outside of a scheduled release.
* **priority:high** - Should be implemented in the next scheduled release if possible.
* **priority:low** - Needs to be implemented but it is not on a specific schedule.
* **priority:maybe-someday** - A nice-to-have feature that we want to track, but may never get time to implement.


#### Effort
There are three options for effort, which is an estimate of how work-intensive the change will be.

* **effort:low** - The change is straightforward and would not be difficult to implement.
* **effort:high** - An entirely new macro, or a fix or enhancement that requires many code changes or multiple dependencies.
* **effort:no-clue** - The change requires some research to determine how the fix would be implemented before we can make a estimate. Once we have enough information to make an estimate, this label should be changed to *effort:high* or *effort:low*.


#### Type
There are four options for type, which indicates what sort of issue this is.

* **type:bug** - The issue is in regard to a macro displaying unexpected and unintended behavior caused by a problem in the code.
* **type:defect** - The issue is in regard to behavior that is as intended, but which produces an undesirable result, perhaps due to a use that was unaccounted for in the original design.
* **type:enhancement** - The issue is a request for a new feature to be added to an already-existing macro.
* **type:new-macro** - The issue is a request for an entirely new macro to be developed.

#### Misc
Assorted labels to be used as needed.

* **duplicate** - The issue is the same as another open or closed issue. Both should be given this label.


### Prioritizing issues
The labels make it relatively easy to identify which issues should be tackled first. Each type of label has an order of priority (as listed above), and the types themselves take priority over each other as listed above. If you sort by labels and pick the highest priority label for each type that has issues attached to it, you'll get a more manageable list of which issues require attention.

For example, all *priority:high* issues should be completed before any *priority:low* issues are started; within the *priority:high* list, all *effort:low* issues should be completed before any *effort:high* issues; and within the *effort:low* list, all *type:bug* issues should be completed before any *type:defect* issues.
