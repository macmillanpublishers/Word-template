# General strategy
Errors are handled within each class in an `ErrorChecker` method that takes the VBA `Err` object as an argument and sorts through the `Err.Number` property with a `Select` block. Note that `Err.Number` is a Long data type, and error numbers 0 to 512 are reserved for built-in system errors. Custom errors can be numbered 513 to 65,535. When users need to be notified of the error, this is done through the `ErrorChecker` method.

Our custom error numbers start at 20000, to avoid conflicts with the custom error numbers used in the [VBA-tools](https://github.com/VBA-tools) modules, which start at 10000. The second significant digit (i.e., second from the left) indicates the class module the error is listed in. Thus, all custom errors for the `MacFile_` class start at 20000, errors for the `Paths_` class start at 21000, and so on. If we end up with more than 10 classes, we'll increment the left-most digit as well.

Custom error numbers are assigned via a `Private Enum` at the top of each module. The item's name should be the name of the procedure that produced it, followed by "Err" and an integer that increments for that member (i.e. the second custom error for the `MacFile_.Download` method is `DownloadErr2 = 20007`.

We may move to error handling to a single class in the future.

# Custom error numbers and descriptions
| Error Number | Class | Enum Name | Description |
| ------------ | ----- | ------ | ----------- |
| 20000 |  | `err_MacErrGeneral` | Generic custom error. May not need, but there it is. |
| 20001 | `MacFile_` | `err_GroupNameInvalid` | The value entered for the `GroupName` property is not present in the `config.json` file, and therefore the `Property Let` procedure failed. The `GroupName` must match one of the objects in the `"files"` object. |
| 20002 | `MacFile_` | `err_GroupNameNotSet` | The `Property Get` procedure was attempted before the `Property Let` procedure had been executed. Using the `AssignFile` method to create new `MacFile_` objects should solve this problem. |
| 20003 | `MacFile_` | `err_SpecificFileInvalid` | The value entered for the `SpecificFile` property is not present in the `config.json` file, and therefore the `Property Let` procedure failed. The `SpecificFile` must match one of the objects in the `GroupName` object. |
| 20004 | `MacFile_` | `err_SpecificFileNotSet` | The `Property Get` procedure was attempted before the `Property Let` procedure had been executed. Using the `AssignFile` method to create new `MacFile_` objects should solve this problem. |
| 20005 | `MacFile_` | `err_DeleteThisDoc` | Attempted to delete file that was currently executing code, so `Delete` was aborted. |
| 20006 | `MacFile_` | `err_TempDeleteFail` | Deletion of previous file in `FullTempPath` failed. Download aborted. |
| 20007 | `MacFile_` | `err_NoInternet` | No network connection. Download aborted. |
| 20008 | `MacFile_` | `err_Http404` | File returned HTTP status of 404, not found. Check if file is correctly posted to DownloadURL. |
| 20009 | `MacFile_` | `err_BadHttpStatus` | File returned bad HTTP status other than 404. Check log for actual status code and description. |
| 20010 | `MacFile_` | `err_DownloadFail` | File download failed. |
| 20011 | `MacFile_` | `err_LocalDeleteFail` | File in final install location could not be deleted. If it was because the file was open, the user was notified. |
| 20012 | `MacFile_` | `err_LocalCopyFail` | Everything else worked, but the file did not end up in the final directory. |
| 20013 | `MacFile_` | `err_LocalReadOnly` | Final dir for file is read-only. |
| 20014 | `MacFile_` | `err_TempReadOnly` | Temp path is read-only. |
| 20015 | `Paths_` | `err_TempMissing` | Temp path does not exist. |
| 20016 | `SharedMacros_` | `err_FileNotThere` | `IsItThere()` function returned false. Check log to see which procedure raised the error--does not get raised by the function itself. |
| 20017 | `SharedMacros_` | `err_NotWordFormat` | `IsWordFormat()` function returned false. Check log to see which procedure raised the error--does not get raised in the `IsWordFormat()` function itself. |
| 20018 | `Paths_` | `err_ConfigPathNull` | There is no "FullConfigPath" custom document property set in the template. |
| 20019 | `Paths_` | `err_RootDirInvalid` | The value for the root directory in the `config.json` file is not an option in the `RootDir` property. |
| 20020 | `Paths_` | `err_RootDirMissing` | The directory returned by the RootDir property doesn't exist. |
