# Security settings


# Mac won't load template in Startup
On OS X, because we download the template files from the internet, they are stripped of their filetype codes. This prevents Word from recognizing that the file is a template and causes Word to silently fail to load the Add-In. Additionally, Apple's `quarantine` XML attribute is added to the file. 

The following two Terminal commands will fix both issues (though change the path if needed).

```
$ xattr -d com.apple.quarantine /Applications/Microsoft\ Office\ 
  2011/Office/Startup/Word/GtUpdater.dotm
 ```
 
 ``` 
$ xattr -wx com.apple.FinderInfo "57 58 54 4D 4D 53 57 44 00 10 00 00 
  00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" 
  /Applications/Microsoft\ Office\ 2011/Office/Startup/Word/GtUpdater.dotm
```

Verify that the `xattribute` is properly set with this command:

``` 
xattr -l /Applications/Microsoft\ Office\ 2011/Office/Startup/Word/GtUpdater.dotm
```

Results should look like:

![good to go](images/confirm-xattr.png)