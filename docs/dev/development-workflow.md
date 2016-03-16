# Version control
Intro about Git, links to some helpful info.


# Branches
[Based on this model, w/ one tweak.](http://nvie.com/posts/a-successful-git-branching-model/)

## master

## releases
a.k.a. beta testing

## develop

## features

## hotfix

# Detailed steps for merging branches
More complicated than usual because of binary files.

## No merge conflicts
Still should probably checkout binaries, import all modules, add CustomUI.xml or whatever, test, add, commit, push.

Is there a way to be sure that there are no changes that won't be in the template files when closing pull request?


# Deployment (H2 below merging steps?)
## Notifications
### Westchester
email:  APS@antares.co.in, Macmillan_PreEdit@wbrt.com, Tina.Mingolello@wbrt.com, Terry.Colosimo@wbrt.com
24-48 hours before release for bug fixes
5-7 business days before release for template and/or macro updates, especially if they involve process changes 

### Users
Send notification via MailChimp to local users (use Word Styles Newsletter segment) with basic info about the updates (24-48 hours before deploying new release).

### MacSupport
