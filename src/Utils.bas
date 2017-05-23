Attribute VB_Name = "Utils"
' =============================================================================
'         UTILS
' =============================================================================

' ***** USE *******************************************************************
' Procedures that can be used in multiple VBA projects. Trying to be as general
' purpose as possible.

' ***** DEPENDENCIES **********************************************************
' Must be imported with other submodules: VBA-Dictionary, VBA-JSON


' *****************************************************************************
'     DECLARATIONS
' *****************************************************************************
Option Explicit



' *****************************************************************************
'     PROCEDURES
' *****************************************************************************


' ===== DocFileCopy ===========================================================
' Copies a doc from Origin to Destination. Makes sure files are closed before
' copying, and reopens docs if they were open to start. DestinationFile will
' be overwritten. Displays VBE if it was open before.

' PARAMS
' OriginFile[String]: full path to file to copy from
' DestinationFile[String]: full path to file to copy to

Public Sub DocFileCopy(OriginFile As String, DestinationFile As String)
  Dim blnOriginOpened As Boolean
  Dim blnDestinationOpened As Boolean
  Dim blnVbEditorOpen As Boolean

' If this code though, don't try to copy while code is running
  If OriginFile = ThisDocument.FullName Or DestinationFile = ThisDocument.FullName Then
    Exit Sub
  End If

' Save and close relevant docs, record if were open
  blnOriginOpened = Utils.DocSaveClose(OriginFile)
  blnDestinationOpened = Utils.DocSaveClose(DestinationFile)
  blnVbEditorOpen = Application.VBE.MainWindow.Visible

' Check if loaded as an Add-In, and if so unload it
  Dim blnOriginAddin As Boolean
  Dim blnDestAddin As Boolean
  Dim strOriginFileAndExt As String
  Dim strDestFileAndExt As String
  
  strOriginFileAndExt = Utils.GetFileName(OriginFile)
  strDestFileAndExt = Utils.GetFileName(DestinationFile)
  
  blnOriginAddin = Utils.UnloadAddIn(strOriginFileAndExt)
  blnDestAddin = Utils.UnloadAddIn(strDestFileAndExt)
  
  VBA.FileCopy Source:=OriginFile, Destination:=DestinationFile
  
  If blnOriginOpened = True Then
    Documents.Open (OriginFile)
  End If
  
  If blnDestinationOpened = True Then
    Documents.Open (DestinationFile)
  End If
  
  If blnOriginAddin = True Then
    Utils.LoadAddIn AddinName:=strOriginFileAndExt, DisableAutoMacros:=True
  End If
  
  If blnDestAddin = True Then
    Utils.LoadAddIn AddinName:=strDestFileAndExt, DisableAutoMacros:=True
  End If
  
  Application.VBE.MainWindow.Visible = blnVbEditorOpen
End Sub


' ===== DocSaveClose ==========================================================
' Makes sure document is saved and closed, including various validation things.

' PARAMS
' Path[String]: Full path to file w/ filename and ext

' RETURNS
' True = Doc was originally open
' False = Doc was originally closed

Public Function DocSaveClose(Path As String) As Boolean
  Dim blnFileOpen As Boolean
  blnFileOpen = Utils.IsOpen(Path)
  
' But if it's this code here don't close it!
  If Path = ThisDocument.FullName Then
    Documents(Path).Save
  Else
    If blnFileOpen = True Then
      Documents(Path).Close SaveChanges:=wdSaveChanges
    End If
  End If
  
  DocSaveClose = blnFileOpen

End Function


' ===== DocOpenSave ===========================================================
' Makes sure document is opened and saved.

' PARAMS
' Path[String]: Full path to file with filename and ext

' RETURNS
' True = Doc was originally open
' False = Doc was originally closed

Public Function DocOpenSave(Path As String) As Boolean
  Dim blnFileOpen As Boolean
  blnFileOpen = Utils.IsOpen(Path)
  
  If blnFileOpen = False Then
    Documents.Open Path
  End If
  
  DocOpenSave = blnFileOpen

End Function

' ===== GetFileExtension =======================================================
' Returns file extension WITHOUT dot. If no dot, returns null string.

Public Function GetFileExtension(File As String) As String
  Dim lngExtLen As Long
  lngExtLen = InStr(StrReverse(File), ".") - 1
  
  If lngExtLen > 0 Then
    GetFileExtension = Right(File, lngExtLen)
  Else
    GetFileExtension = vbNullString
  End If
End Function

' ===== GetFileNameOnly ========================================================
' Strips file extension from end and path from beginning of string.

Public Function GetFileNameOnly(File As String) As String
  Dim lngLastSeparatorPosition As Long
  Dim lngExtensionDotLen As Long
  Dim lngFileNameStart As Long
  Dim lngFileNameLength As Long
  
  lngLastSeparatorPosition = InStrRev(File, Application.PathSeparator)
  lngExtensionDotLen = InStr(StrReverse(File), ".")
  lngFileNameStart = lngLastSeparatorPosition + 1
  lngFileNameLength = Len(File) - lngLastSeparatorPosition - lngExtensionDotLen
  GetFileNameOnly = Mid(File, lngFileNameStart, lngFileNameLength)
End Function

' ===== GetFileName ===========================================================
' Strips path and returns file name with extension.

Public Function GetFileName(File As String) As String
  GetFileName = Right(File, InStr(StrReverse(File), _
    Application.PathSeparator) - 1)
End Function

' ===== GetPath ================================================================
' Returns path to parent directory from string full name to file.

' PARAMS
' FullName: Full path to a file

' RETURNS
' Sting path to parent directory, with no trailing separator

Public Function GetPath(FullName As String) As String
' Opposite of GetFileName, which returns null if not a file
  Dim strFileName As String
  strFileName = GetFileName(FullName)
  
  Dim strPath As String
  If strFileName = vbNullString Then
    strPath = FullName
  Else
    strPath = Left(FullName, Len(FullName) - Len(strFileName))
  End If
  
' Remove trailing path separator
  If Right(strPath, 1) = Application.PathSeparator Then
    strPath = Left(strPath, Len(strPath) - 1)
  End If
  
  GetPath = strPath

End Function

' ===== DebugPrint =============================================================
' Use instead of `DebugPrint`. Print to Immediate Window AND write to a file.
' Immediate Window has a small buffer and isn't very useful if you are debugging
' something that ends up crashing the app.

' Actual `DebugPrint` can take more complex arguments but here we'll just take
' anything that can evaluate to a string.

' Need to set "VbaDebug" environment variable to True also

Public Sub DebugPrint(Optional StringExpression As Variant)

  If Environ("VbaDebug") = True Then
  ' First just standard Debug.Print:
  ' Get the string we'll write
    Dim strMessage As String
    strMessage = Now & ": " & StringExpression
    Debug.Print strMessage
  
  ' Second, write to file
  ' Create file name
  ' !!! activeDoc.Path sometimes writes to STARTUP dir. Also if running
  ' with Folder Actions (like Validator), new file in dir will error
  ' How to write to a static location?
    Dim strOutputFile As String
    strOutputFile = Environ("USERPROFILE") & Application.PathSeparator & _
      "Desktop" & Application.PathSeparator & "immediate_window.txt"
  
    Dim FileNum As Integer
    FileNum = FreeFile ' next file number
    Open strOutputFile For Append As #FileNum
    Print #FileNum, strMessage
    Close #FileNum ' close the file
  End If
 
End Sub

' ===== IsOldMac ==============================================================
' Checks this is a Mac running Office 2011 or earlier. Good for things like
' checking if we need to account for file paths > 3 char (which 2011 can't
' handle but Mac 2016 can.

Public Function IsOldMac() As Boolean
  IsOldMac = False
  #If Mac Then
      If Application.Version < 16 Then
          IsOldMac = True
      End If
  #End If
End Function

' ===== DocPropExists =========================================================
' Tests if a particular custom document property exists in the document. If
' it's already a Document object we already know that it exists and is open
' so we don't need to test for those here. Should be tested somewhere in
' calling procedure though.

Public Function DocPropExists(objDoc As Document, PropName As String) As Boolean
  DocPropExists = False

' Note DocumentProperties returns a Collection
  Dim docProps As DocumentProperties
  Set docProps = objDoc.CustomDocumentProperties

  Dim varProp As Variant

  If docProps.Count > 0 Then
      For Each varProp In docProps
          If varProp.Name = PropName Then
              DocPropExists = True
              Exit Function
          End If
      Next varProp
  Else
      DocPropExists = False
  End If
End Function

' ===== IsOpen ================================================================
' Tests if the Word document is currently open.

Public Function IsOpen(DocPath As String) As Boolean
  Dim objDoc As Document
  IsOpen = False
  If IsItThere(DocPath) = True Then
    If IsWordFormat(DocPath) = True Then
      If Documents.Count > 0 Then
        For Each objDoc In Documents
          If objDoc.FullName = DocPath Then
            IsOpen = True
            Exit Function
          End If
        Next objDoc
      End If
    End If
  End If
End Function

' ===== IsWordFormat ==========================================================
' Checks extension to see if file is a Word document or template. Notably,
' does not test if it's a file type that Word CAN open (e.g., .html), just
' if it's a native Word file type.

' Ignores final character for newer file types, just checks for .dot / .doc

Public Function IsWordFormat(PathToFile As String) As Boolean
  Dim strExt As String
  strExt = Left(Right(PathToFile, InStr(StrReverse(PathToFile), ".")), 4)
  If strExt = ".dot" Or strExt = ".doc" Then
    IsWordFormat = True
  Else
    IsWordFormat = False
  End If
End Function

' ===== IsLocked ==============================================================
' Tests if any file is locked by some kind of process.

Public Function IsLocked(FilePath As String) As Boolean
  On Error GoTo IsLockedError
  IsLocked = False
  If IsItThere(FilePath) = False Then
    Exit Function
  Else
    Dim FileNum As Long
    FileNum = FreeFile()
    ' If the file is already in use, next line will raise an error:
    ' "70: Permission denied" (file is open, Word doc is loaded as add-in)
    ' "75: Path/File access error" (File is read-only, etc.)
    Open FilePath For Binary Access Read Write Lock Read Write As FileNum
    Close FileNum
  End If
IsLockedFinish:
  Exit Function
    
IsLockedError:
  If Err.Number = 70 Or Err.Number = 75 Then
      IsLocked = True
  End If
End Function

' ===== IsItThere =============================================================
' Check if file or directory exists on PC or Mac.
' Dir() doesn't work on Mac 2011 if file is longer than 32 char

Public Function IsItThere(Path As String) As Boolean
  'Remove trailing path separator from dir if it's there
  If Right(Path, 1) = Application.PathSeparator Then
    Path = Left(Path, Len(Path) - 1)
  End If

  If IsOldMac = True Then
    Dim strScript As String
    strScript = "tell application " & Chr(34) & "System Events" & Chr(34) & _
        "to return exists disk item (" & Chr(34) & Path & Chr(34) _
        & " as string)"
    IsItThere = ShellAndWaitMac(strScript)
  Else
    Dim strCheckDir As String
    strCheckDir = Dir(Path, vbDirectory)
    
    If strCheckDir = vbNullString Then
        IsItThere = False
    Else
        IsItThere = True
    End If
  End If
End Function

' ===== ParentDirExists =======================================================
' If `FilePath` is the full path to a file (that may or may not exist), then
' this checks that the directory the file is in exists. Good for checking paths
' to files before you create them.

Public Function ParentDirExists(FilePath As String) As Boolean
  Dim strDir As String
  Dim strFile As String
  Dim lngSep As Long
  ParentDirExists = False
  ' Separate directory from file name
  lngSep = InStrRev(FilePath, Application.PathSeparator)
  
  If lngSep > 0 Then
    strDir = VBA.Left(FilePath, lngSep - 1)  ' NO trailing separator
    strFile = VBA.Right(FilePath, Len(FilePath) - lngSep)
'    DebugPrint strDir & " | " & strFile

    ' Verify file name string is in fact plausibly a file name
    If InStr(strFile, ".") > 0 Then
      ' NOW we can check if the directory exists:
      ParentDirExists = IsItThere(strDir)
      Exit Function
    End If
  End If
End Function

' ===== KillAll ===============================================================
' Deletes file (or folder?) on PC or Mac. Mac can't use Kill() if file name
' is longer than 32 char. Returns true if successful.
    
Public Function KillAll(Path As String) As Boolean
  On Error GoTo KillAllError
  If IsItThere(Path) = True Then
    ' Can't delete file if it's installed as an add-in
    If IsInstalledAddIn(Path) = True Then
        AddIns(Path).Installed = False
    End If
    ' Mac 2011 can't handle file paths > 32 char
    #If Mac Then
      If Application.Version < 16 Then
        Dim strCommand As String
        strCommand = MacScript("return quoted form of posix path of " & Path)
        strCommand = "rm " & strCommand
        ShellAndWaitMac (strCommand)
      Else
        Kill (Path)
      End If
    #Else
      Kill (Path)
    #End If

    ' Make sure it worked
    If IsItThere(Path) = False Then
      KillAll = True
    Else
      KillAll = False
    End If
  Else
    KillAll = True
  End If
KillAllFinish:
  Exit Function
    
KillAllError:
  Dim strErrMsg As String
  Select Case Err.Number
    Case 70     ' File is open
      strErrMsg = "Please close all other Word documents and try again."
      MsgBox strErrMsg, vbCritical, "Macmillan Tools Error"
      KillAll = False
      Resume KillAllFinish
    Case Else
      Exit Function
  End Select
End Function

' ===== IsInstalledAddInError =================================================
' Check if the file is currently loaded as an AddIn. Because we can't delete
' it if it is loaded (though we can delete it if it's just referenced but
' not loaded).

Public Function IsInstalledAddIn(FileName As String) As Boolean
  Dim objAddIn As AddIn
  For Each objAddIn In AddIns
    ' Check if in collection first; throws error if try to check .Installed
    ' but it's not even referenced.
    If objAddIn.Name = FileName Then
      If objAddIn.Installed = True Then
        IsInstalledAddIn = True
      Else
        IsInstalledAddIn = False
      End If
      Exit For
    End If
  Next objAddIn
End Function


' ===== UnloadAddIn ===========================================================
' Checks if file is loaded as a Word add-in, and if it is, it unloads it.

' ASSUMPTIONS
' Addin name is same as file name plus extension

' PARAMS
' AddinName[String]: usually AddinName is file name with extension, but it can
'   be other things so be careful

' RETURNS
' True = file was loaded as an addin
' False = file was NOT loaded as an addin

Public Function UnloadAddIn(AddinName As String) As Boolean
  If Utils.IsInstalledAddIn(AddinName) = False Then
    UnloadAddIn = False
  Else
    Application.AddIns(AddinName).Installed = False
    UnloadAddIn = True
  End If
End Function

' ===== LoadAddin =============================================================
' If file isn't already INSTALLED as an add-in (not just part of Addins), this
' installed/loads it. Writes a file

' ASSUMPTIONS
' You've already checked that the file is in the AddIns collection, but haven't
' checked that it's "Installed" (i.e. loaded)

' PARAMS
' AddinName[String]: usually AddinName is file name with extension, but it can
'   be other things so be careful
' DisableAutoExec[Boolean]: If True, writes a text file to same path as addin.
'   The AutoExec macro needs to check for existance of this file, and quit if found

' NOTES
' The AddIn.Installed property doesn't actually test if it is part of the AddIns
' collection. It sets whether the AddIn is "loaded" to run or not. The file must
' already be in the AddIns collection or you'll get an error when you try to
' get that property.

Public Sub LoadAddIn(AddinName As String, DisableAutoMacros As Boolean)
  If Utils.IsInstalledAddIn(AddinName) = False Then
    Dim objAddIn As AddIn
    Set objAddIn = Application.AddIns(AddinName)

  ' If we need to disable AutoExec macros, write a file to signal that procedure
    If DisableAutoMacros = True Then
    ' Get path where addin file is saved
      Dim strDisableFlagFile As String
      strDisableFlagFile = objAddIn.Path & Application.PathSeparator & _
        "DISABLE_AUTO_EXEC.txt"
      
    ' Write a file there (content doesn't really matter)

      Utils.OverwriteTextFile TextFile:=strDisableFlagFile, NewText:="True"
    End If
    
  ' Load the add-in
    objAddIn.Installed = True
    
  ' Delete that file if we wrote it earlier
    If DisableAutoMacros = True Then
      Utils.KillAll Path:=strDisableFlagFile
    End If
  
  End If
End Sub


' ===== DisableAutoExec =======================================================
' Checks if the "DISABLE_AUTO_EXEC.txt" file written in LoadAddIn is present in
' same dir as the file executing the code.

' PARAMS
' TemplatePath[String]: Path to directory the template file you are trying to load
' is in, with no trailing separator.

' ASSUMPTIONS
' You use Utils.LoadAddIn to load an addin, and indicate whether you want to
' disable autoexec macros or not.
' You check this function at the start of any AutoExec macros.

' RETURNS: Boolean
' True: file is present, exit AutoExec
' False: file is not present, do not exit

Public Function DisableAutoExec(TemplatePath As String) As Boolean
  Dim strFlagFilePath As String
  strFlagFilePath = TemplatePath & Application.PathSeparator & "DISABLE_AUTO_EXEC.txt"
  DisableAutoExec = Utils.IsItThere(strFlagFilePath)
End Function



' ===== ShellAndWaitMac =======================================================
' Sends shell command to AppleScript on Mac (to replace missing functions!)

Public Function ShellAndWaitMac(Cmd As String) As String
  Dim result As String
  Dim scriptCmd As String ' Macscript command
  #If Mac Then
    scriptCmd = "do shell script " & Chr(34) & Cmd & Chr(34) & Chr(34)
    result = MacScript(scriptCmd) ' result contains stdout, should you care
    'DebugPrint result
    ShellAndWaitMac = result
  #End If
End Function

' ===== OverwriteTextFile =====================================================
' Pretty self explanatory. TextFile parameter should be full path.

Public Sub OverwriteTextFile(TextFile As String, NewText As String)
  Dim FileNum As Integer
  ' Will create file if not exist, but parent dir must exist
  If ParentDirExists(TextFile) = True Then
    FileNum = FreeFile ' next file number
    Open TextFile For Output Access Write As #FileNum
    Print #FileNum, NewText ' overwrite information in the text of the file
    Close #FileNum ' close the file
  Else
    ' directory is invalid

  End If
End Sub

' ===== AppendTextFile ========================================================
' Appends Contents string to file that already exists.

Public Sub AppendTextFile(TextFile As String, Contents As String)
' TextFile should be full path
  On Error GoTo AppendTextFileError
  Dim FileNum As Integer
' Will create file if not exist, but parent dir must exist
  TextFile = VBA.Replace(TextFile, "/", Application.PathSeparator)
  If ParentDirExists(TextFile) = True Then
    FileNum = FreeFile ' next file number
    Open TextFile For Append As #FileNum
    Print #FileNum, Contents
    Close #FileNum ' close the file
  Else
    ' directory is invalid
  End If
End Sub


' ===== IsReadOnly ============================================================
' Tests if the file or directory is read-only -- does NOT test if file exists,
' because sometimes you'll need to do that before this anyway to do something
' different.

' Mac 2011 can't deal with file paths > 32 char
    
Public Function IsReadOnly(Path As String) As Boolean

    If IsOldMac() = True Then
        Dim strScript As String
        Dim blnWritable As Boolean
        
        strScript = _
            "set p to POSIX path of " & Chr(34) & Path & Chr(34) & Chr(13) & _
            "try" & Chr(13) & _
            vbTab & "do shell script " & Chr(34) & "test -w \" & Chr(34) & _
            "$(dirname " & Chr(34) & " & quoted form of p & " & Chr(34) & _
            ")\" & Chr(34) & Chr(34) & Chr(13) & _
            vbTab & "return true" & Chr(13) & _
            "on error" & Chr(13) & _
            vbTab & "return false" & Chr(13) & _
            "end try"
            
        blnWritable = MacScript(strScript)
        
        If blnWritable = True Then
            IsReadOnly = False
        Else
            IsReadOnly = True
        End If
    Else
        If (GetAttr(Path) And vbReadOnly) <> 0 Then
            IsReadOnly = True
        Else
            IsReadOnly = False
        End If
    End If
End Function

' ===== ReadTextFile ==========================================================

Public Function ReadTextFile(Path As String, Optional FirstLineOnly As Boolean _
  = True) As String

' load string from text file

    Dim fnum As Long
    Dim strTextWeWant As String
    
    fnum = FreeFile()
    Open Path For Input As fnum
    
    If FirstLineOnly = False Then
        strTextWeWant = Input$(LOF(fnum), fnum)
    Else
        Line Input #fnum, strTextWeWant
    End If
    
    Close fnum
    
    ReadTextFile = strTextWeWant

End Function



' ===== ExistsInCollection ====================================================
' Tests if item is in the collection by trying to access the item; if it is not
' in the collection, it generates an error.

' Params:
' TestCollection: Collection object
' TestItem: Object or whatever to test

Public Function ExistsInCollection(TestCollection As Collection, TestItem As _
  Variant) As Boolean

  If IsObject(TestItem) = True Then
    Dim objItem As Object
    For Each objItem In TestCollection
      If objItem Is TestItem Then
        ExistsInCollection = True
        Exit Function
      End If
    Next
  Else
    Dim varItem As Variant
    For Each varItem In TestCollection
      If varItem = TestItem Then
        ExistsInCollection = True
        Exit Function
      End If
    Next
  End If
End Function

' ===== SortCollection ========================================================
' No natice Collection.Sort method, so we'll convert to an array, sort that,
' then convert back.

' PARAMS
' SourceCollection: ByRef so actual collection will be changed in calling
' procedure.

' TODO: If this adds too much overhead, research implementations of sorting
' algorithms on the Collection directly (info exists for VBA, just not that
' necessary right now).

Public Sub SortCollection(ByRef SourceCollection As Collection)
  Dim arr_TempArray() As Variant
  arr_TempArray = ToArray(SourceCollection)
  WordBasic.SortArray arr_TempArray
  Set SourceCollection = ToCollection(arr_TempArray)
End Sub

' ===== ToArray ===============================================================
' Converts a Collection to an array.

' PARAMS:
' SourceCollection: A Collection object to be converted

' TODO
' Add option to remove Empty collection items when converting.

Public Function ToArray(SourceCollection As Collection) As Variant()
' lngIndexAdjust so it doesn't matter if array is base 0, 1, etc.
  Dim arr() As Variant
  Dim lngIndexAdjust As Long
  Dim lngUBound As Long

' Test array here to get default lower bound, -1 to adjust UBound
  ReDim arr(2)
  lngIndexAdjust = LBound(arr) - 1
  lngUBound = SourceCollection.Count + lngIndexAdjust

  ReDim arr(lngUBound)
  Dim varItem As Variant
  
  For Each varItem In SourceCollection
    lngIndexAdjust = lngIndexAdjust + 1
    arr(lngIndexAdjust) = varItem
  Next varItem

  ToArray = arr

End Function

' ===== ToCollection ==========================================================
' Convert an array to a Collection

Public Function ToCollection(SourceArray() As Variant) As Collection
  Dim coll As Collection
  Set coll = New Collection
  Dim A As Long

  For A = LBound(SourceArray) To UBound(SourceArray)
    coll.Add SourceArray(A)
  Next A

  Set ToCollection = coll
End Function
