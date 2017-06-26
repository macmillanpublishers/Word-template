Attribute VB_Name = "MacroHelpers"

' All should be declared as Public for use from other modules

' *****************************************************************************
'           DECLARATIONS
' *****************************************************************************
Option Explicit
Private Const strModule As String = "MacroHelpers."
Public lngErrorCount As Long


' assign to actual document we're working on
' to do: probably better managed via a class
Public activeDoc As Document

' *****************************************************************************
'           ERROR HANDLING STUFF
' *****************************************************************************
' Technically class methods (which are listed farther down) but it's easier to
' maintain close to the error number enum, which has to appear above other
' procedures. May move to its own class one day.

' ===== MacErrors =========================================================
' Each error we raise anywhere in this class should be included here. Note that
' error numbers 0 - 512 are reserved for system errors. Include property or
' method name in enum just to be clear for debugging.

Public Enum MacError
    err_MacErrGeneral = 20000       ' Not sure if we'll need, but here it is.
    err_GroupNameInvalid = 20001
    err_GroupNameNotSet = 20002
    err_SpecificFileInvalid = 20003
    err_SpecificFileNotSet = 20004
    err_DeleteThisDoc = 20005
    err_TempDeleteFail = 20006
    err_NoInternet = 20007
    err_Http404 = 20008
    err_BadHttpStatus = 20009
    err_DownloadFail = 20010
    err_LocalDeleteFail = 20011
    err_LocalCopyFail = 20012
    err_LocalReadOnly = 20013
    err_TempReadOnly = 20014
    err_TempMissing = 20015
    err_FileNotThere = 20016
    err_NotWordFormat = 20017
    err_ConfigPathNull = 20018
    err_RootDirInvalid = 20019
    err_RootDirMissing = 20020
    err_LogReadOnly = 20021
    err_DirectoryMissing = 20022
    err_ParaIndexInvalid = 20023
    err_BacktickCharFound = 20024
    err_DocProtectionOn = 20025
    err_NotArray = 20026
End Enum

' ===== ErrorChecker ==========================================================
' Send all errors here, to make tracking/maintaining error handling easier.
' Be sure to add the error to the MacErrors enum.

' DO NOT use any other of our functions in this function, because those need to
' direct errors here and using them here as well could create an infinite loop.
'
' USE: ErrorChecker() returns False if error is handled; procedure can continue

Public Function ErrorChecker(objError As Object, Optional strValue As _
    String) As Boolean
' For debugging
  Stop
  
    ' strValue - varies based on type of error passed. use for things like
    ' file name, path, whatever is being checked by that errored.
  lngErrorCount = lngErrorCount + 1
  DebugPrint "ErrorChecker " & lngErrorCount & vbNewLine & _
    "(" & objError.Source & ") " & objError.Number & ":" & vbNewLine _
    & objError.Description
  
  If lngErrorCount > 15 Then
    DebugPrint "ERROR LOOP STOPPED"
    End
  End If
  
  ' New On Error statement RESETS the Err object, so get our values out before
  ' we set the ErrorChecker for this procedure.
  Dim lngErrNumber As Long
  Dim strErrDescription As String
  Dim strErrSource As String
  
  lngErrNumber = objError.Number
  strErrDescription = objError.Description  ' For system errors
  strErrSource = objError.Source
  
'  On Error GoTo ErrorCheckerError

  Dim strErrMessage As String
  Dim blnNotifyUser As Boolean
  Dim strHelpContact As String

  ' ----- Set defaults --------------------------------------------------------
  ' Only need to change in Select statement below if want to set either to
  ' False.
  blnNotifyUser = True
  ErrorChecker = True
  strHelpContact = vbNewLine & vbNewLine & "Email workflows@macmillan.com if" _
    & " you need help."

  ' ----- Check if parameter was passed ---------------------------------------
  If strValue = vbNullString Then
    strValue = "UNKNOWN"
  End If

  ' ----- Check for errors ----------------------------------------------------
  ' Make sure we actually have an error, cuz you never know.
  If lngErrNumber <> 0 Then

  ' ----- Handle specific errors ----------------------------------------------
  ' Each case should be a MacError enum. Even if we aren't doing anything
  ' to fix the error, still enter the following:
  '   strErrDescription = "Description of the error for the log."
  '   strErrMessage = "Message for the user if we're notifying them."
  Select Case lngErrNumber
  
    Case 5941, 5834
    ' 5941: Item not present in collection
    ' 5834: Item with the specified name does not exist
    ' Most common cause of these is referencing a style that doesn't exist.
    
    ' BUT: How can we be sure this means *style* is not present?
    
      'Anyway, If style is not present, add style
      Dim myStyle As Style
      Dim styleType As WdStyleType
      If InStr(strValue, "span") > 0 Then
        styleType = wdStyleTypeCharacter
      Else
        styleType = wdStyleTypeParagraphOnly
      End If
      Set myStyle = activeDoc.Styles.Add(strValue, styleType)
      ErrorChecker = False
      DebugPrint "ErrorChecker: False"
      Exit Function
    ' List all built-in errors we want to trap for before general sys error line
    Case 91 ' Object variable or With block variable not set.
      ' May be caused if `activeDoc` global var is not set
      If activeDoc Is Nothing Then
        Set activeDoc = ActiveDocument
        ErrorChecker = False
        DebugPrint "ErrorChecker: False"
        Exit Function
      End If
      
    Case Is < 513
      ' Errors 0 to 512 are system errors
      strErrMessage = "Something unexpected happened. Please click OK to exit." _
        & vbNewLine & "Value: " & strValue
    Case MacError.err_GroupNameInvalid
      strErrDescription = "Invalid value for GroupName property: " & strValue
      strErrMessage = "The value you've entered for the GroupName property, " _
        & "is not valid. Make sure you only use file groups that are in the " _
        & "config.json file."
    Case MacError.err_GroupNameNotSet
      strErrDescription = "GroupName property has not been Let."
      strErrMessage = "You can't Get the GroupName property before it has " _
        & "been Let. Try the MacFile_.AssignFile method to create a new object in this class."
    Case MacError.err_SpecificFileInvalid
      strErrDescription = "Invalid value for SpecificFile property: " _
          & strValue
      strErrMessage = "The you've entered an invalid value for the" _
          & " SpecificFile property. Make sure you only use " & _
          "specific file types that are in the config.json file."
    Case MacError.err_SpecificFileNotSet
      strErrDescription = "SpecificFile property has not been Let."
      strErrMessage = "You can't Get the SpecificFile property " & _
          "before it has been Let. Try the MacFile_.AssignFile " & _
          "method to create a new object in this class."
    Case MacError.err_DeleteThisDoc
      strErrDescription = "Can't delete file that is currently " & _
          "executing code: " & strValue
      strErrMessage = "The file you are trying to delete is " & _
          "currently executing macro code."
    Case MacError.err_TempDeleteFail
      strErrDescription = "Failed to delete the previous file in the " _
          & "temp directory: " & strValue
      strErrMessage = "We can't download the file; a temp file " & _
          "is still there."
    Case MacError.err_NoInternet
      strErrDescription = "No network connection. Download aborted."
      strErrMessage = "We weren't able to download the file " & _
          "because we can't connect to the internet. Check your " & _
          "network connection and try again."
    Case MacError.err_Http404
      strErrDescription = "File HTTP status 404. Check if DownloadURL" _
          & " is correct, and file is posted: " & strValue
      strErrMessage = "Could not download file from the internet."
    Case MacError.err_BadHttpStatus
      strErrDescription = "File HTTP status: " & strValue & _
            ". Download aborted."
      strErrMessage = "There is some problem with the file you are" _
            & " trying to download."
        ' Need to get Source as passed in object first, so do this last
    Case MacError.err_DownloadFail
      strErrDescription = "File download failed: " & strValue
      strErrMessage = "Download failed."
    Case MacError.err_LocalDeleteFail
        ' Utils.KillAll() will notify user if file is open
      strErrDescription = "File in final install location could not be " & _
        "deleted. If the file was open, the user was notified: " & strValue
      blnNotifyUser = False
    Case MacError.err_LocalCopyFail
      strErrDescription = "File not saved to final directory: " & strValue
      strErrMessage = "There was an error installing the Macmillan template."
    Case MacError.err_LocalReadOnly
      strErrDescription = "Final dir for file is read-only: " & strValue
      strErrMessage = "The folder you are trying to access is read-only."
    Case MacError.err_TempReadOnly
      strErrDescription = "Temp dir is read-only: " & strValue
      strErrMessage = "Your temp folder is read-only."
    Case MacError.err_TempMissing
      strErrDescription = "Temp directory is missing: " & strValue
       strErrMessage = "There is an error with your temp folder."
    Case MacError.err_FileNotThere
       strErrDescription = "File does not exist: " & strValue
       strErrMessage = "The file " & strValue & " does " _
            & "not exist."
    Case MacError.err_NotWordFormat
      strErrDescription = "File extension is not a native Word " & _
            "document or template: " & strValue
       strErrMessage = "This file does not appear to be a Word " & _
            "file: " & strValue
    Case MacError.err_ConfigPathNull
       strErrDescription = "FullConfigPath custom doc property is not " _
            & "set in the document."
      strErrMessage = "We can't find the config.json file because " _
            & "the local path is not in the template properties."
    Case MacError.err_RootDirInvalid
      strErrDescription = "Value for root directory in config.json is" _
            & " not an option in the RootDir property: " & strValue
      strErrMessage = "The folder where we save the Tools template" _
          & " doesn't exist."
    Case MacError.err_LogReadOnly
      strErrDescription = "Log file is read only: " & strValue
      strErrMessage = "There is a problem with the logs."
    Case MacError.err_DirectoryMissing
      strErrDescription = "The directory " & strValue & " is missing."
      strErrMessage = strErrDescription
    Case MacError.err_ParaIndexInvalid
      strErrDescription = "The requested paragraph is out of range."
      strErrMessage = strErrDescription
    Case MacError.err_BacktickCharFound
      strErrDescription = "Backtick (`) character found in manuscript. A " & _
        "macro was probably run before and failed."
      strErrMessage = strErrDescription
    Case MacError.err_DocProtectionOn
      strErrDescription = "Document protection is enabled. Ask original user" _
        & " to unlock the file and try again."
      strErrMessage = strErrDescription
    Case MacError.err_NotArray
      strErrDescription = "Variable is not an array."
      strErrMessage = strErrDescription
    Case Else
      strErrDescription = "Undocumented error - " & strErrDescription
      strErrMessage = "Not sure what's going on here."
  End Select

  Else
      strErrDescription = "Everything's A-OK. Why are you even reading this?"
      blnNotifyUser = False
      ErrorChecker = False
  End If

  ' ----- WRITE ERROR LOG ---------------------------------------------------
  ' Output text file with error info, user could send via email.
  ' Do not use WriteToLog function, because that sends errors here as well.

  Dim strErrMsg As String
  Dim LogFileNum As Long
  Dim strTimeStamp As String
  Dim strErrLog As String
  Dim strFileName As String

' Check activeDoc:
  If activeDoc Is Nothing Then
    Set activeDoc = ActiveDocument
  End If

  ' Write error log to same location as current file.
  ' Format date so it can be part of file name. Only including date b/c users
  ' will likely run things repeatedly before asking for help, and don't want
  ' to generate a bunch of files if include time as well.
  strFileName = Replace(Right(ThisDocument.Name, InStrRev(activeDoc.Name, _
    ".") - 1), " ", "")
  strErrLog = activeDoc.Path & Application.PathSeparator & "ALERT_" & _
    strFileName & "_" & Format(Date, "yyyy-mm-dd") & ".txt"
'    DebugPrint strErrLog
  ' build error message, including timestamp
  strErrMsg = Format(Time, "hh:mm:ss - ") & strErrSource & vbNewLine & _
      lngErrNumber & ": " & strErrDescription & vbNewLine
  LogFileNum = FreeFile ' next file number
  Open strErrLog For Append As #LogFileNum ' creates the file if doesn't exist
  Print #LogFileNum, strErrMsg ' write information to end of the text file
  Close #LogFileNum ' close the file
  
  ' Do not display alerts for Bookmaker project (automated)
  If WT_Settings.InstallType = "server" Then
    blnNotifyUser = False
  End If
  
  If blnNotifyUser = True Then
      strErrMessage = strErrMessage & vbNewLine & vbNewLine & strHelpContact
      MsgBox Prompt:=strErrMessage, Buttons:=vbExclamation, Title:= _
          "Macmillan Tools Error"
  End If
  DebugPrint "ErrorChecker: " & ErrorChecker
  Exit Function

ErrorCheckerError:
  ' Important note: Recursive error checking is perhaps a bad idea -- if the
  ' same error gets triggered, procedure will get called too many times and
  ' cause an "out of stack space" error and crash.
  DebugPrint Err.Number & ": " & Err.Description
  ErrorChecker = True
End Function

' ===== GlobalCleanup =========================================================
' A variety of resetting/cleanup functions

Sub GlobalCleanup()
  On Error GoTo GlobalCleanupError
  zz_clearFind
  If Not activeDoc Is Nothing Then
    Set activeDoc = Nothing
  End If
  Application.DisplayAlerts = wdAlertsAll
  Application.ScreenUpdating = True
  Application.ScreenRefresh
  On Error GoTo 0

GlobalCleanupError:
  ' Halts ALL execution, resets all variables, unloads all userforms, etc.
  End
End Sub

' ===== WriteToLog ============================================================
' Writes line to log for the file. LogMessage only needs text, timestamp will
' be added in this method.

Public Sub WriteToLog(LogMessage As String, Optional LogFilePath As String)
    On Error GoTo WriteToLogError
    Dim strLogFile As String
    Dim strLogMessage As String
    Dim FileNum As Integer

    ' If no specific path was passed, write to generic log file
    If LogFilePath = vbNullString Then
        strLogFile = Paths_.LogsDir & Application.PathSeparator & _
            "manuscript-tools.log"
    Else
        strLogFile = LogFilePath
    End If

    If IsItThere(strLogFile) = True Then
        If IsReadOnly(strLogFile) = True Then
            Err.Raise MacError.err_LogReadOnly
        End If
    End If

    ' prepend current date and time to message
    strLogMessage = Now & " -- " & LogMessage
    FileNum = FreeFile ' next file number
    Open strLogFile For Append As #FileNum ' creates the file if doesn't exist
    Print #FileNum, strLogMessage ' write information to end of the text file
    Close #FileNum ' close the file
WriteToLogFinish:
    Exit Sub

WriteToLogError:
    Err.Source = Err.Source & strModule & "WriteToLog"
    If MacroHelpers.ErrorChecker(Err, strLogFile) = False Then
        Resume
    Else
        Call MacroHelpers.GlobalCleanup
    End If
End Sub

Public Function IsStyleInUse(StyleName As String) As Boolean
  On Error GoTo IsStyleInUseError
  
' First confirm style is even in document to begin with
  If MacroHelpers.IsStyleInDoc(StyleName) = False Then
    IsStyleInUse = False
    Exit Function
  End If

'  ' If we need to do a Selection.Find use
'  Selection.HomeKey Unit:=wdStory
  Call MacroHelpers.zz_clearFind
  With activeDoc.Range.Find
    .Text = ""
    .Format = True
    .Style = activeDoc.Styles(StyleName)
    .Execute
    
    If .Found = True Then
      IsStyleInUse = True
    Else
      IsStyleInUse = False
    End If
  End With
  
  Exit Function
IsStyleInUseError:
  Err.Source = strModule & "IsStyleInUse"
  If ErrorChecker(Err, StyleName) = False Then
    Resume
  Else
    Call MacroHelpers.GlobalCleanup
  End If
End Function


Public Function IsStyleInDoc(StyleName As String) As Boolean
  On Error GoTo IsStyleInDocError
  Dim blnResult As Boolean: blnResult = True
  Dim TestStyle As Style
  
' Try to access this style. If not present in doc, will error
  Set TestStyle = activeDoc.Styles.Item(StyleName)
  IsStyleInDoc = blnResult
  Exit Function
  
IsStyleInDocError:
' 5941 = "The requested member of the collection does not exist."
' Have to test here, ErrorChecker tries to create style if missing
  If Err.Number = 5941 Then
    blnResult = False
    Resume Next
  End If
' Otherwise, usual error stuff:
  Err.Source = strModule & "IsStyleInDoc"
  If ErrorChecker(Err, StyleName) = False Then
    Resume
  Else
    Call MacroHelpers.GlobalCleanup
  End If
End Function

' ===== SetPathSeparator ======================================================
' Replaces original path separators in string with current file system separators

Public Function SetPathSeparator(strOrigPath As String) As String
' Must pass full path, throws error if no path separators found.
  On Error GoTo SetPathSeparatorError
  Dim strFinalPath As String
  strFinalPath = strOrigPath
  
  Dim strCharactersCollection As Collection
  Dim strCharacter As String
  strCharactersCollection.Add = ":"
  strCharactersCollection.Add = "/"
  strCharactersCollection.Add = "\"
  
  For Each strCharacter In strCharactersCollection
    If InStr(strOrigPath, strCharacter) > 0 Then
      strFinalPath = VBA.Replace(strOrigPath, strOrigPath, _
        Application.PathSeparator)
    End If
  Next strCharacter
  
  SetPathSeparator = strFinalPath
  Exit Function
  
SetPathSeparatorError:
  Err.Source = strModule & "SetPathSeparator"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call MacroHelpers.GlobalCleanup
  End If
End Function


' ===== ParaIndex =============================================================
' Returns the paragraph index of the current selection. Default is to return the
' END paragraph index if selection is more than 1 paragraph. `UseEnd:=False`
' would return the index of the START paragraph.
Public Function ParaIndex(Optional UseEnd As Boolean = True) As Long
  On Error GoTo ParaIndexError
  If UseEnd = True Then
    ParaIndex = activeDoc.Range(0, Selection.End).Paragraphs.Count
  Else
    ParaIndex = activeDoc.Range(0, Selection.Start).Paragraphs.Count
  End If
  Exit Function
ParaIndexError:
  Err.Source = strModule & "ParaIndex"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    MacroHelpers.GlobalCleanup
  End If
End Function

' ===== ParaInfo ==============================================================
' In general: get a variety of info about a paragraph (or its document)
' Most common usage: InfoType = wdActiveEndAdjustedPageNumber

Public Function ParaInfo(paraInd As Long, InfoType As WdInformation) _
  As Variant
  On Error GoTo ParaInfoError
  
 
' Make sure our paragraph index is in range
  If paraInd <= activeDoc.Paragraphs.Count Then
  ' Set range for our paragraph
    Dim rngPara As Range
    Set rngPara = activeDoc.Paragraphs(paraInd).Range
    ParaInfo = rngPara.Information(InfoType)
  Else
    Err.Raise MacError.err_ParaIndexInvalid
  End If

  Exit Function
ParaInfoError:
  Err.Source = strModule & "ParaInfo"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call GlobalCleanup
  End If
End Function

Public Sub zz_clearFind(Optional TargetRange As Range)
  On Error GoTo zz_clearFindError

' If we didn't pass a Range, reset the Selection.Find
' Can sometimes have sticky properties if not reset properly.
  If TargetRange Is Nothing Then
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Wrap = wdFindStop
        .Format = False
        .Forward = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
' Not 100% sure, but I think some Find/Replace issues are caused b/c
' Execute doesn't work (or at least is wonky) if you don't set the
' Replace parameter each time.
        .Execute Replace:=wdReplaceNone
    End With
  Else
    With TargetRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Wrap = wdFindStop
        .Format = False
        .Forward = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceNone
    End With
  End If

'' So we can see updates in the Advanced Find UI
'  If WT_Settings.DebugOn = True Then
'    Application.ScreenRefresh
'  End If
  
  Exit Sub
zz_clearFindError:
' Can't do any replace if doc is password protected, but this runs
' as part of cleanup so need to handle that here:
  If Err.Number = 9099 Then ' "Command is not available"
    Exit Sub
  End If
  Err.Source = strModule & "zz_clearFind"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call GlobalCleanup
  End If
End Sub

Public Function ActiveStories() As Collection
  On Error GoTo ActiveStoriesError
    '------------check for endnotes and footnotes--------------------------
    Dim colStories As Collection
    Set colStories = New Collection

    Dim storyAdd As WdStoryType
    storyAdd = wdMainTextStory
    colStories.Add storyAdd
    
    If activeDoc.Endnotes.Count > 0 Then
        storyAdd = wdEndnotesStory
        colStories.Add storyAdd
    End If
    
    If activeDoc.Footnotes.Count > 0 Then
        storyAdd = wdFootnotesStory
        colStories.Add storyAdd
    End If
    
    Set ActiveStories = colStories
  Exit Function
ActiveStoriesError:
  Err.Source = strModule & "ActiveStories"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call GlobalCleanup
  End If
End Function

Function PatternMatch(SearchPattern As String, SearchText As String, WholeString As Boolean) As Boolean
    ' "SearchPattern" uses Word Find pattern matching, which is not the same as regular expressions
    ' But the RegEx library breaks Word Mac 2011, so we'll do it this way
    ' This is a good reference: http://www.gmayor.com/replace_using_wildcards.htm
    ' "SearchText" is the string you're looking in
    ' "WholeString" is True if you are trying to match the whole string; if just part
    ' of the string is an acceptable match, set to False
        
    ' Need to paste string into a Word doc to use Find pattern matching
    Dim newDoc As New Document
    Set newDoc = Documents.Add(Visible:=False)
    newDoc.Select
    
    Selection.InsertBefore (SearchText)
    ' Insertion point has to be at start of doc for Selection.Find
    Selection.Collapse (wdCollapseStart)
    
    With Selection.Find
        .ClearFormatting
        .Text = SearchPattern
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWholeWord = False
        .MatchCase = True
        .MatchWildcards = True
        .MatchSoundsLike = False
        .Execute
    End With
    
    If Selection.Find.Found = True Then
        If WholeString = True Then
            ' The final paragraph return is the only character the new doc had it in,
            ' it 's not part of the added string
            If InStrRev(Selection.Text, Chr(13)) = Len(Selection.Text) Then
                Selection.MoveEnd Unit:=wdCharacter, Count:=-1
            End If
            
            ' the SearchText requires vbCrLf to start text on a new line, but Word for some reason
            ' strips out the Lf when content is pasted in. CrLf counts as 2 characters but Cr is only
            ' 1, so to get these to match we need to add 1 character to the selection for each line.
            Dim lngLines As Long
            lngLines = activeDoc.ComputeStatistics(wdStatisticLines)
            
            If Len(Selection.Text) + lngLines = Len(SearchText) Then
                PatternMatch = True
            Else
                PatternMatch = False
            End If
        Else
            PatternMatch = True
        End If
    Else
        PatternMatch = False
    End If
    
    newDoc.Close wdDoNotSaveChanges
    
End Function


Public Function EndnotesExist() As Boolean
    Dim StoryRange As Range
    
    EndnotesExist = False
    
    For Each StoryRange In activeDoc.StoryRanges
        If StoryRange.StoryType = wdEndnotesStory Then
            EndnotesExist = True
            Exit For
        End If
    Next StoryRange
End Function

Public Function FootnotesExist() As Boolean
    Dim StoryRange As Range
    
    FootnotesExist = False
    
    For Each StoryRange In activeDoc.StoryRanges
        If StoryRange.StoryType = wdFootnotesStory Then
            FootnotesExist = True
            Exit For
        End If
    Next StoryRange
    
End Function


Function IsArrayEmpty(arr As Variant) As Boolean


    Dim LB As Long
    Dim UB As Long
    
    Err.Clear
    On Error Resume Next
    If IsArray(arr) = False Then
        ' we weren't passed an array, return True
        IsArrayEmpty = True
        Exit Function
    End If

    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    UB = UBound(arr, 1)
    If (Err.Number <> 0) Then
        IsArrayEmpty = True
    Else
  On Error GoTo IsArrayEmptyError
        LB = LBound(arr)
        If LB > UB Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If
  Exit Function
IsArrayEmptyError:
  Err.Source = strModule & "IsArrayEmpty"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call GlobalCleanup
  End If
End Function


Public Sub CreateTextFile(strText As String, suffix As String)

    Application.ScreenUpdating = False
    
    'Create report file
    Dim activeRng As Range
    Set activeRng = activeDoc.Range
    Dim activeDocName As String
    Dim activeDocPath As String
    Dim reqReportDoc As String
    Dim reqReportDocAlt As String
    Dim fnum As Integer
    Dim TheOS As String
    Dim strMacTmp As String
    TheOS = System.OperatingSystem
    
    'activeDocName below works for .doc and .docx
    activeDocName = Left(activeDoc.Name, InStrRev(activeDoc.Name, ".do") - 1)
    activeDocPath = Replace(activeDoc.Path, activeDoc.Name, "")
    
    'create text file
    reqReportDoc = activeDocPath & activeDocName & "_" & suffix & ".txt"
    
    ''''for 32 char Mc OS bug- could check if this is Mac OS too < PART 1
    If Not TheOS Like "*Mac*" Then                      'If Len(activeDocName) > 18 Then        (legacy, does not take path into account)
        reqReportDoc = activeDocPath & "\" & activeDocName & "_" & suffix & ".txt"
    Else
        Dim placeholdDocName As String
        placeholdDocName = "filenamePlacehold_Report.txt"
        reqReportDocAlt = reqReportDoc
        strMacTmp = MacScript("path to temporary items as string")
        reqReportDoc = strMacTmp & placeholdDocName
    End If
    '''end ''''for 32 char Mc OS bug part 1
    
    'set and open file for output
    Dim E As Integer
    fnum = FreeFile()
    Open reqReportDoc For Output As fnum
    
        Print #fnum, strText

    Close #fnum
    
    ''''for 32 char Mc OS bug-<PART 2
    If reqReportDocAlt <> "" Then
    Name reqReportDoc As reqReportDocAlt
    End If
    ''''END for 32 char Mac OS bug-<PART 2
    
    '----------------open Report for user once it is complete--------------------------.
    Dim Shex As Object
    
    If Not TheOS Like "*Mac*" Then
       Set Shex = CreateObject("Shell.Application")
       Shex.Open (reqReportDoc)
    Else
        MacScript ("tell application ""TextEdit"" " & vbCr & _
        "open " & """" & reqReportDocAlt & """" & " as alias" & vbCr & _
        "activate" & vbCr & _
        "end tell" & vbCr)
    End If
End Sub

Function GetText(StyleName As String, Optional ReturnArray As Boolean = False) _
  As Variant
  On Error GoTo GetTextError
  
  Dim fCount As Integer
  Dim styleArray() As Variant

  fCount = 0
  
  'Move selection to start of document
  Selection.HomeKey Unit:=wdStory

      MacroHelpers.zz_clearFind
      With Selection.Find
          .Text = ""
          .Replacement.Text = ""
          .Forward = True
          .Wrap = wdFindStop
          .Format = True
          .Style = activeDoc.Styles(StyleName)
          .MatchCase = False
          .MatchWholeWord = False
          .MatchWildcards = False
          .MatchSoundsLike = False
          .MatchAllWordForms = False
      End With
  
  Do While Selection.Find.Execute = True And fCount < 100            'fCount < 100 so we don't get an infinite loop
      fCount = fCount + 1
      
      'If paragraph return exists in selection, don't select last character (the last paragraph retunr)
      If InStr(Selection.Text, Chr(13)) > 0 Then
          Selection.MoveEnd Unit:=wdCharacter, Count:=-1
      End If
      
      'Assign selected text to variable
      ReDim Preserve styleArray(1 To fCount)
      styleArray(fCount) = Selection.Text
      
      'If the next character is a paragraph return, add that to the selection
      'Otherwise the next Find will just select the same text with the paragraph return
      If InStr(StyleName, "span") = 0 Then        'Don't select terminal para mark if char style, sends into an infinite loop
          Selection.MoveEndWhile Cset:=Chr(13), Count:=1
      End If
  Loop
      
  If fCount = 0 Then
      ReDim styleArray(1 To 1)
      styleArray(1) = ""
  End If
  
  If ReturnArray = False Then
    GetText = Join(SourceArray:=styleArray, Delimiter:=vbNewLine)
  Else
    GetText = styleArray
  End If
  Exit Function
  
GetTextError:
  Err.Source = strModule & "GetText"
  If Err.Number = 5941 Or Err.Number = 5834 Then   ' The style is not present in the document
      GetText = ""
  Else
    If ErrorChecker(Err) = False Then
      Resume
    Else
      Call MacroHelpers.GlobalCleanup
    End If
  End If
  
End Function


' ===== StyleReplace ==========================================================
' Replace all instances of a specific paragraph style with a different style.
' If you want to "remove" the style, replace with "Normal" or whatever. Returns
' False if no replacements were made.

Public Function StyleReplace(SearchStyle As String, ReplaceStyle As String) As _
  Boolean
  On Error GoTo StyleReplaceError
  
  MacroHelpers.zz_clearFind
  With activeDoc.Range.Find
    .Format = True
    .Style = SearchStyle
    .Replacement.Style = ReplaceStyle
    .Execute Replace:=wdReplaceAll
    
    If .Found = True Then
      StyleReplace = True
    Else
      StyleReplace = False
    End If
  
  End With
  MacroHelpers.zz_clearFind
  Exit Function
  
StyleReplaceError:
  Err.Source = strModule & "StyleReplace"
  If ErrorChecker(Err, ReplaceStyle) = False Then
    Resume
  Else
    Call MacroHelpers.GlobalCleanup
  End If
End Function


Function StartupSettings(Optional AcceptAll As Boolean = False) As Boolean
  On Error GoTo StartupSettingsError
' records/adjusts/checks settings and stuff before running the rest of the macro
' returns TRUE if some check is bad and we can't run the macro

' activeDoc is global variable to hold our document, so if user clicks a different
' document during execution, won't switch to that doc.
' ALWAYS set to Nothing first to reset for this macro.
' Then only refer to this object, not ActiveDocument directly.

  Set activeDoc = Nothing
  Set activeDoc = ActiveDocument


' check if file has doc protection on, quit function if it does
  If activeDoc.ProtectionType <> wdNoProtection Then
    If WT_Settings.InstallType = "server" Then
      Err.Raise MacError.err_DocProtectionOn
    Else
      MsgBox "Uh oh ... protection is enabled on document '" & activeDoc.Name & "'." & vbNewLine & _
        "Please unprotect the document and run the macro again." & vbNewLine & vbNewLine & _
        "TIP: If you don't know the protection password, try pasting contents of this file into " & _
        "a new file, and run the macro on that.", , "Error 2"
      StartupSettings = True
      Exit Function
    End If
  End If
  
' check if file has been saved (we can assume it was for Validator)
  If WT_Settings.InstallType = "user" Then
    Dim iReply As Integer
    Dim docSaved As Boolean
    docSaved = activeDoc.Saved
  
    If docSaved = False Then
      iReply = MsgBox("Your document '" & activeDoc & "' contains unsaved changes." & vbNewLine & vbNewLine & _
          "Click OK to save your document and run the macro." & vbNewLine & vbNewLine & "Click 'Cancel' to exit.", _
              vbOKCancel, "Error 1")
      If iReply = vbOK Then
        activeDoc.Save
      Else
        StartupSettings = True
        Exit Function
      End If
    End If
  End If

  ' ========== Turn off screen updating ==========
  Application.ScreenUpdating = False

' Section of registry/preferences file to store settings
  Dim strSection As String
  strSection = "MACMILLAN_MACROS"
    
  ' ========== Remove bookmarks ==========
  Dim bkm As Bookmark
  
  For Each bkm In activeDoc.Bookmarks
    bkm.Delete
  Next bkm
    
  ' ========== Save current cursor location in a bookmark ==========
  If WT_Settings.InstallType = "user" Then
    ' Store current story, so we can return to it before selecting bookmark in Cleanup
    System.ProfileString(strSection, "Current_Story") = Selection.StoryType
    ' next line required for Mac to prevent problem where original selection blinked repeatedly when reselected at end
    Selection.Collapse Direction:=wdCollapseStart
    activeDoc.Bookmarks.Add Name:="OriginalInsertionPoint", Range:=Selection.Range
  End If
    
  ' ========== TRACK CHANGES: store current setting, turn off ==========
  System.ProfileString(strSection, "Current_Tracking") = activeDoc.TrackRevisions
  activeDoc.TrackRevisions = False

  ' ========== Check if changes present and offer to accept all ==========
  ' AcceptAll is a parameter passed from the calling procedure. If calling
  ' from Validator, be sure to set it to True.

  If AcceptAll = True Then
      If FixTrackChanges = False Then
          StartupSettings = True
      End If
  End If
  
  ' ========== Remove content controls ==========
  ' Content controls also break character styles and cleanup
  ' They are used by some imprints for frontmatter templates
  ' for editorial, though.
  ' Doesn't work at all for a Mac, so...
  ' NOTE: New version cleans up Cookbook template. Mac way of checking only works
  ' with template version 3+
  Dim strOrigTemplate As String
  Dim strCookbookMsg As String
  #If Mac Then
      Dim objDocProp As DocumentProperty
      For Each objDocProp In activeDoc.CustomDocumentProperties
        If objDocProp.Name = "OriginalTemplate" Then
          If InStr(objDocProp.Value, "CookbookTemplate_v") > 0 Then
            strCookbookMsg = "It looks like you are cleaning up a cookbook manuscript. " & _
              "Note that cleanup specific to Macmillan's Cookbook template only works " & _
              "on Windows PCs. Please ask your PE or another friendly coworker to run " & _
              "this macro for you."
            MsgBox strCookbookMsg
            Exit For
          End If
        End If
      Next objDocProp
  #Else
      ' Run both, 2nd one clears non-recipe content controls
      CleanUpRecipeContentControls
      ClearContentControls
  #End If
  
' ========== Delete field codes ==========
' Fields break cleanup and char styles, so we delete them (but retain their
' result, if any). Furthermore, fields make no sense in a manuscript, so
' even if they didn't break anything we don't want them.
' Note, however, that even though linked endnotes and footnotes are
' types of fields, this loop doesn't affect them.
' NOTE: Moved this to separate procedure to use Matt's code.
' Must run AFTER content control cleanup.

  Dim colStoriesUsed As Collection
  Set colStoriesUsed = MacroHelpers.ActiveStories()
  Call UpdateUnlinkFieldCodes(colStoriesUsed)


  ' ========== STATUS BAR: store current setting and display ==========
  ' Run after Content control cleanup
  If WT_Settings.InstallType = "user" Then
    System.ProfileString(strSection, "Current_Status_Bar") = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
  End If
  
  Exit Function
  
StartupSettingsError:
  Err.Source = strModule & "StartupSettings"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call MacroHelpers.GlobalCleanup
  End If
End Function

' ===== UpdateUnlinkFieldCodes ================================================
' Cycles through all Fields in activeDoc. Updates, unlocks, and unlinks
' each field. If this is our cookbook template with the automatic TOC, that
' will be unlinked as well.

Public Sub UpdateUnlinkFieldCodes(Optional p_stories As Collection)
  Dim objField As Field
  Dim thisRange As Range
  Dim strContent As String
  Dim blnTOCpresent As Boolean
  
  ' Test if we need to run ReMapTOCStyles later
  blnTOCpresent = False
  
  ' p_stories is optional; if an array of stories is not passed,
  ' just use the main text story here
  If p_stories.Count = 0 Then
    Dim mainStory As WdStoryType
    mainStory = wdMainTextStory
    Set p_stories = New Collection
    p_stories.Add mainStory
  End If

  Dim varStory As Variant
  Dim currentStory As WdStoryType

  For Each varStory In p_stories
    currentStory = varStory
    Set thisRange = activeDoc.StoryRanges(currentStory)
    If thisRange.Fields.Count > 0 Then
      For Each objField In thisRange.Fields
'            DebugPrint thisRange.Fields.Count
        With objField
          If .Type = wdFieldTOC Then
            blnTOCpresent = True
          End If
          
          .Update
          .Locked = False
          .Unlink
        
        End With
      Next objField
    End If
  Next
  
  ' If automatic TOC was unlinked above, need to map built-in TOC styles to ours
  If blnTOCpresent = True Then
    Call ReMapTOCStyles
  End If

End Sub


' ===== ReMapTOCStyles =========================================================
' Replaces built-in TOC styles with Macmillan equivalents (based on Dict)

Private Sub ReMapTOCStyles()
  On Error GoTo ReMapTOCStylesError
  
  Dim objStyleMapDict As Dictionary
  Dim objDictKey As Variant
  Dim objDictValue As Variant
  Dim rngActiveDoc As Range
  Dim myStyle As Style  ' for error handling
  
  Set objStyleMapDict = CookbookTOCStyleMap
  Set rngActiveDoc = activeDoc.Range
   
  For Each objDictKey In objStyleMapDict.Keys()
  
    'Need to add a check if style is present in Document &/or in use
    objDictValue = objStyleMapDict(objDictKey)
    
    Call zz_clearFind
    With rngActiveDoc.Find
      .Text = ""
      .Replacement.Text = ""
      .Wrap = wdFindContinue
      .Format = True
      .Style = objDictKey
      .Replacement.Style = objDictValue
      .Execute Replace:=wdReplaceAll
    End With
  Next
  
  Exit Sub

ReMapTOCStylesError:
  If Err.Number = 5834 Or Err.Number = 5941 Then  ' style not present
    Set myStyle = activeDoc.Styles.Add(Name:=objDictKey, _
      Type:=wdStyleTypeParagraph)
    Resume
  Else
    If WT_Settings.InstallType = "user" Then
      MsgBox "Oops, something happened! Email workflows@macmillan.com and " & _
        "let them know that something's wrong." & vbNewLine & vbNewLine & _
        "Error " & Err.Number & ": " & Err.Description
    Else
      ErrorChecker Err
    End If
  End If
End Sub



' ===== CookbookTOCStyleMap ====================================================
'
' Hardcoded style-map for TOC styles (for cookbook template)

Private Function CookbookTOCStyleMap()
  Dim objStyleMapDict As New Dictionary

  objStyleMapDict.Add "TOC 1", "TOC Frontmatter Head (cfmh)"
  objStyleMapDict.Add "TOC 2", "TOC Backmatter Head (cbmh)"
  objStyleMapDict.Add "TOC 3", "TOC Part Number  (cpn)"
  objStyleMapDict.Add "TOC 4", "TOC Part Title (cpt)"
  objStyleMapDict.Add "TOC 5", "TOC Chapter Number (ccn)"
  objStyleMapDict.Add "TOC 6", "TOC Chapter Title (cct)"
  objStyleMapDict.Add "TOC 7", "TOC Author (cau)"
  objStyleMapDict.Add "TOC 8", "TOC Level-1 Chapter Head (ch1)"
  objStyleMapDict.Add "TOC 9", "TOC Chapter Subtitle (ccst)"
  
  Set CookbookTOCStyleMap = objStyleMapDict
End Function

Private Function FixTrackChanges() As Boolean
' returns True if changes were fixed or not present, False if changes remain in doc
  On Error GoTo FixTrackChangesError
    Dim N As Long
    Dim oComments As Comments
    Set oComments = activeDoc.Comments
    
    FixTrackChanges = True
    
    'See if there are tracked changes or comments in document
    On Error Resume Next
    Selection.HomeKey Unit:=wdStory   'start search at beginning of doc
    'search for a tracked change or comment. error if none are found.
    WordBasic.NextChangeOrComment
    
' If there are changes, ask user if they want macro to accept changes or cancel
    If Err.Number = 0 Then
      If WT_Settings.InstallType = "user" Then
        If MsgBox("Bookmaker doesn't like comments or tracked changes, but it appears that you have some in your document." _
          & vbCr & vbCr & "Click OK to ACCEPT ALL CHANGES and DELETE ALL COMMENTS right now and continue with the Bookmaker Requirements Check." _
          & vbCr & vbCr & "Click CANCEL to stop the Bookmaker Requirements Check and deal with the tracked changes and comments on your own.", _
          273, "Are those tracked changes I see?") = vbCancel Then           '273 = vbOkCancel(1) + vbCritical(16) + vbDefaultButton2(256)
              FixTrackChanges = False
              Exit Function
        Else 'User clicked OK, so accept all tracked changes and delete all comments
          activeDoc.AcceptAllRevisions
          For N = oComments.Count To 1 Step -1
              oComments(N).Delete
          Next N
          Set oComments = Nothing
        End If
      End If
    Else
      FixTrackChanges = True
    End If
    Exit Function
    
FixTrackChangesError:
  Err.Source = strModule & "FixTrackChanges"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call MacroHelpers.GlobalCleanup
  End If
    
End Function

' ===== CleanUpRecipeContentControls ===========================================
'
' For cleaning up Cookstr Cookbook templates:
'   1. Calls "UpdateUnlinkTOC" sub to update, unlock, and unlink TOC fields
'   2. Cycles through all Content Controls in doc, Unlocks each CC.
'   3. For all CCs of type 'group', if nested CCs are empty + have tag cookbook
'       it deletes the group & all contents. If not, it deletes the group CC
'       preserving contents)
'   4. For all non-Group CC's, if CC is in paragraphs styled as Design Note:
'       if CC is the last/only Content Control in the DN para, the para is deleted
'   5. For any "Edirotial" CC's with placeholder content, the CC range text is
'       set to match placeholder content (and persists when CC is deleted).
'   6. All other ContentControls, CC is deleted, preserving non-placeholder content

Private Sub CleanUpRecipeContentControls()
  Dim objCC As ContentControl
  Dim objCCs As ContentControls
  Dim objGroupCC As ContentControl
  Dim rngCC As Range
  Dim rngGroupCC As Range
  Dim rngIndexPara As Range
  Dim lngCCsInPara As Long
  Dim lngEmptyCCinGroup As Long
  Dim lngParaIndex As Long
  Dim objStyleMapDict As Dictionary
  Set objCCs = activeDoc.ContentControls
  Set objStyleMapDict = CookbookTOCStyleMap
  
  For Each objCC In objCCs
    If objCC.LockContentControl = True Then
        objCC.LockContentControl = False
    End If
    If objCC.Type = 7 Then            'check for grouped CC's
        Set rngGroupCC = objCC.Range
        lngEmptyCCinGroup = 0
        For Each objGroupCC In rngGroupCC.ContentControls
            If objGroupCC.PlaceholderText.Value = objGroupCC.Range.Text And _
                objGroupCC.Tag = "cookbook" Then
                lngEmptyCCinGroup = lngEmptyCCinGroup + 1
            End If
        Next
        If lngEmptyCCinGroup = rngGroupCC.ContentControls.Count Then
            DebugPrint "Deleting a blank '" & _
                rngGroupCC.ContentControls(1).Title & "' CC group"
            objCC.Delete True
        Else
            objCC.Delete False
        End If
    ElseIf objCC.Tag = "cookbook" Or objCC.Tag = "cookbooks" Or _
        objCC.Title = "Pub Year" Then
        Set rngCC = objCC.Range
        If rngCC.ParagraphStyle = "Design Note (dn)" Then
            DebugPrint "Deleting a Design Note para with ContentControl: " _
                & objCC.Title
            lngParaIndex = activeDoc.Range(0, rngCC.End).Paragraphs.Count
            Set rngIndexPara = activeDoc.Paragraphs(lngParaIndex).Range
            lngCCsInPara = rngIndexPara.ContentControls.Count
            DebugPrint lngCCsInPara & "is the lngpcount"
            If rngIndexPara.ContentControls(lngCCsInPara).ID = objCC.ID Then
                'to verify this is the last Content Control in this para
                activeDoc.Paragraphs(lngParaIndex).Range.Delete
            End If
        Else
            If objCC.Title = "Editorial" And objCC.Range.Text = _
                objCC.PlaceholderText.Value Then
                objCC.Range.Text = objCC.PlaceholderText.Value
                DebugPrint "Setting blank 'Editorial' CCs to placeholder txt"
            End If
            objCC.Delete False
            DebugPrint "Deleting CC (preserving content) from para: " & _
                rngCC.ParagraphStyle
        End If
    Else ' It's not a cookbook control and we want to remove but leave content
      objCC.Delete False
    End If
  Next

End Sub

Private Sub ClearContentControls()
' Run CleanupRecipeContentControls first.
  On Error GoTo ClearContentControlsError
    'This is it's own sub because doesn't exist in Mac Word,
    ' breaks whole sub if included
    Dim cc As ContentControl
    
    For Each cc In activeDoc.ContentControls
        cc.Delete
    Next
  Exit Sub
ClearContentControlsError:
  Err.Source = strModule & "ClearContentControls"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call MacroHelpers.GlobalCleanup
  End If
End Sub



Sub Cleanup()
  On Error GoTo CleanUpError
    ' resets everything from StartupSettings sub.
    Dim cleanupDoc As Document
    Set cleanupDoc = activeDoc
    
    ' Section of registry/preferences file to get settings from
    Dim strSection As String
    strSection = "MACMILLAN_MACROS"
    
    ' restore Status Bar to original setting
    ' If key doesn't exist, set to True as default
    Dim currentStatus As String
    currentStatus = System.ProfileString(strSection, "Current_Status_Bar")
    
    If currentStatus <> vbNullString Then
        Application.StatusBar = currentStatus
    Else
        Application.StatusBar = True
    End If
    
    ' reset original Track Changes setting
    ' If key doesn't exist, set to false as default
    Dim currentTracking As String
    currentTracking = System.ProfileString(strSection, "Current_Tracking")
    
    If currentTracking <> vbNullString Then
        cleanupDoc.TrackRevisions = currentTracking
    Else
        cleanupDoc.TrackRevisions = False
    End If
    
    ' return to original cursor position
    ' If key doesn't exist, search in main doc
    Dim currentStory As WdStoryType
    currentStory = System.ProfileString(strSection, "Current_Story")
    
    If cleanupDoc.Bookmarks.Exists("OriginalInsertionPoint") = True Then
        If currentStory = 0 Then
            cleanupDoc.StoryRanges(currentStory).Select
        Else
            cleanupDoc.StoryRanges(wdMainTextStory).Select
        End If
        
        Selection.GoTo what:=wdGoToBookmark, Name:="OriginalInsertionPoint"
        cleanupDoc.Bookmarks("OriginalInsertionPoint").Delete
    End If
    
    ' Turn Screen Updating on and refresh screen
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    Exit Sub
CleanUpError:
  If WT_Settings.InstallType = "server" Then
    Err.Source = strModule & "CleanUp"
    If ErrorChecker(Err) = False Then
      Resume
    Else
      Call MacroHelpers.GlobalCleanup
    End If
  End If
End Sub

Function HiddenTextSucks(StoryType As Long) As Boolean
'v. 3.1 patch : redid this whole thing as an array, addedsmart quotes, wrap toggle var
  On Error GoTo HiddenTextSucksError
'    DebugPrint StoryType
    Dim activeRng As Range
    Set activeRng = activeDoc.StoryRanges(StoryType)
    ' No, really, it does. Why is that even an option?
    ' Seriously, this just deletes all hidden text, based on the
    ' assumption that if it's hidden, you don't want it.
    ' returns a Boolean in case we want to notify user at some point
    
    HiddenTextSucks = False
    
    ' If Hidden text isn't shown, it won't be deleted, which
    ' defeats the purpose of doing this at all.
    Dim blnCurrentHiddenView As Boolean
    blnCurrentHiddenView = activeDoc.ActiveWindow.View.ShowAll
    activeDoc.ActiveWindow.View.ShowAll = True

    
    Dim aCounter As Long
    aCounter = 0
    
    ' Select whole doc (story, actually)
    activeRng.Select

    With Selection.Find
        .ClearFormatting
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Font.Hidden = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute ReplaceWith:="", Replace:=wdReplaceAll
    End With
    
    Do While Selection.Find.Execute = True And aCounter < 500
        'aCounter < 500 so we don't get an infinite loop
        aCounter = aCounter + 1
        HiddenTextSucks = True
    Loop
    
    ' Now restore Hidden Text view settings
    activeDoc.ActiveWindow.View.ShowAll = blnCurrentHiddenView
    Exit Function
HiddenTextSucksError:
    Err.Source = strModule & "HiddenTextSucks"
    If ErrorChecker(Err) = False Then
        Resume
    Else
        MacroHelpers.GlobalCleanup
    End If
    
End Function


Sub ClearPilcrowFormat(StoryType As WdStoryType)
 On Error GoTo ClearPilcrowFormatError
' A pilcrow is the paragraph mark symbol. This clears all formatting and styles from
' pilcrows as found via ^p
    ' Change to story ranges?
    Dim activeRange As Range
    Set activeRange = activeDoc.StoryRanges(StoryType)

    With activeRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^13"       ' need to use ^13 if using wildcards
        .Replacement.Text = "^p"    ' DON'T replace with ^13, removes para style
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Replacement.Style = "Default Paragraph Font"
        .Replacement.Font.Italic = False
        .Replacement.Font.Bold = False
        .Replacement.Font.Underline = wdUnderlineNone
        .Replacement.Font.AllCaps = False
        .Replacement.Font.SmallCaps = False
        .Replacement.Font.StrikeThrough = False
        .Replacement.Font.Subscript = False
        .Replacement.Font.Superscript = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
  Exit Sub
ClearPilcrowFormatError:
  If WT_Settings.InstallType = "server" Then
    Err.Source = strModule & "ClearPilcrowFormat"
    If ErrorChecker(Err) = False Then
      Resume
    Else
      MacroHelpers.GlobalCleanup
    End If
  End If
End Sub

Sub StyleAllHyperlinks(Optional StoriesInUse As Collection)
  On Error GoTo StyleAllHyperlinksError
    ' StoriesInUse is a collection of wdStoryTypes in use
    ' Clears active links and adds macmillan URL char styles
    ' to any proper URLs.
    ' Breaking up into sections because AutoFormat does not apply hyperlinks to FN/EN stories.
    ' Also if you AutoFormat a second time it undoes all of the formatting already applied to hyperlinks
    
    Call zz_clearFind
    
    Dim varStory As Variant
    Dim curStory As WdStoryType
    For Each varStory In StoriesInUse
      curStory = varStory
        'Styles hyperlinks, must be performed after PreserveWhiteSpaceinBrkStylesA
        Call StyleHyperlinksA(StoryType:=curStory)
    Next
    
    Call MacroHelpers.AutoFormatHyperlinks
    
    For Each varStory In StoriesInUse
      curStory = varStory
        Call StyleHyperlinksB(StoryType:=curStory)
    Next
 Exit Sub
StyleAllHyperlinksError:
  Err.Source = strModule & "StyleAllHyperlinks"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    MacroHelpers.GlobalCleanup
  End If
End Sub

Private Sub StyleHyperlinksA(StoryType As WdStoryType)
 On Error GoTo StyleHyperlinksAError
    ' PRIVATE, if you want to style hyperlinks from another module,
    ' call StyleAllHyperlinks sub above.
    ' added by Erica 2014-10-07, v. 3.4
    ' removes all live hyperlinks but leaves hyperlink text intact
    ' then styles all URLs as "span hyperlink (url)" style
    ' -----------------------------------------
    ' this first bit removes all live hyperlinks from document
    ' we want to remove these from urls AND text; will add back to just urls later
    Dim activeRng As Range

    Set activeRng = activeDoc.StoryRanges(StoryType)
    ' remove all embedded hyperlinks regardless of character style
    ' Must use Fields obj not Hyperlink obj because if "empty" hyperlink is in
    ' doc, it will return as part of the Hyperlinks collection but will error
    ' when try to delete or access any properties.
    Dim fld As Field
    If activeRng.Fields.Count > 0 Then
      For Each fld In activeRng.Fields
      ' wdFieldKindNone = invalid field
        If fld.Kind <> wdFieldKindNone And fld.Type = wdFieldHyperlink Then
        ' If field is a link but no text appears in the document for it,
        ' just delete the whole thing (otherwise replace link w/ display text)
          If Len(fld.result.Text) = 0 Then
            fld.Delete
          Else
            fld.Unlink
          End If
        End If
      Next fld
    End If

    '------------------------------------------
    'removes all hyperlink styles
    Dim HyperlinkStyleArray(3) As String
    Dim P As Long
    
    HyperlinkStyleArray(1) = "Hyperlink"        'built-in style applied automatically to links
    HyperlinkStyleArray(2) = "FollowedHyperlink"    'built-in style applied automatically
    HyperlinkStyleArray(3) = "span hyperlink (url)" 'Macmillan template style for links
    
    For P = 1 To UBound(HyperlinkStyleArray())
        With activeRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Style = HyperlinkStyleArray(P)
            .Replacement.Style = activeDoc.Styles("Default Paragraph Font")
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Next
    
    Exit Sub
    
StyleHyperlinksAError:
        '5834 means item does not exist
        '5941 means style not present in collection
        If Err.Number = 5834 Or Err.Number = 5941 Then
            
            'If style is not present, add style
            Dim myStyle As Style
            Set myStyle = activeDoc.Styles.Add(Name:="span hyperlink (url)", Type:=wdStyleTypeCharacter)
            Resume
'            ' Used to add highlight color, but actually if style is missing, it's
'            ' probably a MS w/o Macmillan's styles and the highlight will be annoying.
'            'If missing style was Macmillan built-in style, add character highlighting
'            If myStyle = "span hyperlink (url)" Then
'                activeDoc.Styles("span hyperlink (url)").Font.Shading.BackgroundPatternColor = wdColorPaleBlue
'            End If
        Else
          Err.Source = strModule & "StyleHyperlinksA"
          If MacroHelpers.ErrorChecker(Err) = False Then
            Resume
          Else
            Call MacroHelpers.GlobalCleanup
          End If
        End If

End Sub

Private Sub AutoFormatHyperlinks()
  On Error GoTo AutoFormatHyperlinksError
    ' PRIVATE, if you want to style hyperlinks from another module,
    ' call StyleAllHyperlinks sub above.
    '--------------------------------------------------
    ' converts all URLs to hyperlinks with built-in "Hyperlink" style
    ' because some show up as plain text
    ' Note this also removes all blank paragraphs regardless of style,
    ' so needs to come after sub PreserveWhiteSpaceinBrkA
    
    
    Dim f1 As Boolean, f2 As Boolean, f3 As Boolean
    Dim f4 As Boolean, f5 As Boolean, f6 As Boolean
    Dim f7 As Boolean, f8 As Boolean, f9 As Boolean
    Dim f10 As Boolean, f11 As Boolean
      
    'This first bit autoformats hyperlinks in main text story
    With Options
        ' Save current AutoFormat settings
        f1 = .AutoFormatApplyHeadings
        f2 = .AutoFormatApplyLists
        f3 = .AutoFormatApplyBulletedLists
        f4 = .AutoFormatApplyOtherParas
        f5 = .AutoFormatReplaceQuotes
        f6 = .AutoFormatReplaceSymbols
        f7 = .AutoFormatReplaceOrdinals
        f8 = .AutoFormatReplaceFractions
        f9 = .AutoFormatReplacePlainTextEmphasis
        f10 = .AutoFormatReplaceHyperlinks
        f11 = .AutoFormatDeleteAutoSpaces
        ' Only convert URLs
        .AutoFormatApplyHeadings = False
        .AutoFormatApplyLists = False
        .AutoFormatApplyBulletedLists = False
        .AutoFormatApplyOtherParas = False
        .AutoFormatReplaceQuotes = False
        .AutoFormatReplaceSymbols = False
        .AutoFormatReplaceOrdinals = False
        .AutoFormatReplaceFractions = False
        .AutoFormatReplacePlainTextEmphasis = False
        .AutoFormatReplaceHyperlinks = True
        .AutoFormatDeleteAutoSpaces = False
        ' Perform AutoFormat
        activeDoc.Content.AutoFormat
        ' Restore original AutoFormat settings
        .AutoFormatApplyHeadings = f1
        .AutoFormatApplyLists = f2
        .AutoFormatApplyBulletedLists = f3
        .AutoFormatApplyOtherParas = f4
        .AutoFormatReplaceQuotes = f5
        .AutoFormatReplaceSymbols = f6
        .AutoFormatReplaceOrdinals = f7
        .AutoFormatReplaceFractions = f8
        .AutoFormatReplacePlainTextEmphasis = f9
        .AutoFormatReplaceHyperlinks = f10
        .AutoFormatDeleteAutoSpaces = f11
    End With
    
    'This bit autoformats hyperlinks in endnotes and footnotes
    ' from http://www.vbaexpress.com/forum/showthread.php?52466-applying-hyperlink-styles-in-footnotes-and-endnotes
    Dim oTemp As Document
    Dim oNote As Range
    Dim oRng As Range

    Set oTemp = Documents.Add(Visible:=False)
    
    If activeDoc.Footnotes.Count >= 1 Then
        Dim oFN As Footnote
        For Each oFN In activeDoc.Footnotes
            Set oNote = oFN.Range
            Set oRng = oTemp.Range
            oRng.FormattedText = oNote.FormattedText
            'oRng.Style= "Footnote Text"
            Options.AutoFormatReplaceHyperlinks = True
            oRng.AutoFormat
            oRng.End = oRng.End - 1
            oNote.FormattedText = oRng.FormattedText
        Next oFN
        Set oFN = Nothing
    End If
    
    If activeDoc.Endnotes.Count >= 1 Then
        Dim oEN As Endnote
        For Each oEN In activeDoc.Endnotes
            Set oNote = oEN.Range
            Set oRng = oTemp.Range
            oRng.FormattedText = oNote.FormattedText
            'oRng.Style= "Endnote Text"
            Options.AutoFormatReplaceHyperlinks = True
            oRng.AutoFormat
            oRng.End = oRng.End - 1
            oNote.FormattedText = oRng.FormattedText
        Next oEN
        Set oEN = Nothing
    End If
    
    oTemp.Close SaveChanges:=wdDoNotSaveChanges
    Set oTemp = Nothing
    Set oRng = Nothing
    Set oNote = Nothing
  Exit Sub
  
AutoFormatHyperlinksError:
  Err.Source = strModule & "AutoFormatHyperlinks"
  If MacroHelpers.ErrorChecker(Err) = False Then
    Resume
  Else
    Call MacroHelpers.GlobalCleanup
  End If
End Sub

Private Sub StyleHyperlinksB(StoryType As WdStoryType)
  On Error GoTo StyleHyperlinksBError
    ' PRIVATE, if you want to style hyperlinks from another module,
    ' call StyleAllHyperlinks sub above.
    '--------------------------------------------------
    ' apply macmillan URL style to hyperlinks we just tagged in Autoformat
    Dim activeRng As Range
    Set activeRng = activeDoc.StoryRanges(StoryType)
    With activeRng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = "Hyperlink"
        .Replacement.Style = activeDoc.Styles("span hyperlink (url)")
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' -----------------------------------------------
    ' Removes all hyperlinks from the document (that were added with AutoFormat)
    ' Text to display is left intact, macmillan style is left intact
    With activeRng
        While .Hyperlinks.Count > 0
            .Hyperlinks(1).Delete
        Wend
    End With
  Exit Sub
StyleHyperlinksBError:
  Err.Source = strModule & "StyleHyperlinksB"
  If MacroHelpers.ErrorChecker(Err) = False Then
    Resume
  Else
    Call MacroHelpers.GlobalCleanup
  End If
End Sub


' ===== IsNewLine =============================================================
' Returns True if the string parameter contains a new line and ONLY a new line.
' Making a separate function to test all kinds of new lines, since sometimes
' files contain mixed line ending characters. Also note that vbNewLine constant
' returns different on Win vs. Mac 2011 vs. Mac 2016

Public Function IsNewLine(strValue As String) As Boolean
  On Error GoTo IsNewLineError
  If strValue = vbCr Or strValue = vbLf Or strValue = vbCr & vbLf Then
    IsNewLine = True
  End If
  Exit Function
  
IsNewLineError:
  Err.Source = strModule & "IsNewLine"
  If MacroHelpers.ErrorChecker(Err) = False Then
    Resume
  Else
    Call MacroHelpers.GlobalCleanup
  End If
End Function


' ===== GetTextByIndex ========================================================
' Returns text of paragraph with index number, with trailing newlines and other
' whitespace removed.

Public Function GetTextByIndex(ParaIndex As Long) As String
  Dim strParaText As String
  strParaText = ActiveDocument.Paragraphs(ParaIndex).Range.Text
  If IsNewLine(Right(strParaText, 1)) = True Then
    strParaText = Left(strParaText, Len(strParaText) - 1)
  End If
  Debug.Print Trim(strParaText)
  GetTextByIndex = Trim(strParaText)
End Function


' ===== PageBreakCleanup ======================================================
' Remove any page break characters or styles that are previous sibling of a
' section start style.

' Tags any remaining page break characters with the page break style.

' TODO: Figure out how to handle section breaks

Public Sub PageBreakCleanup()
' Add paragraph breaks around every page break character (so we know for sure
' paragraph style of break won't apply to any body text). Will add extra blank
' paragraphs that we can clean up later.

  Dim currentRng As Range
  Set currentRng = activeDoc.Range

  MacroHelpers.zz_clearFind currentRng
  With currentRng.Find
    .Text = "^m"
    .Replacement.Text = "^p^m^p"
    .Format = True
    .Replacement.Style = WT_StyleConfig.MacmillanStyles("pagebreak")
    .Execute Replace:=wdReplaceAll
  End With

' If we had an unstyled page break char, new trailing ^p is wrong style
' Use this to make sure all correct style.
  MacroHelpers.zz_clearFind currentRng
  With currentRng.Find
    .Text = "^m^13{1,}"
    .Replacement.Text = "^m^p"
    .Format = True
    .MatchWildcards = True
    .Replacement.Style = WT_StyleConfig.MacmillanStyles("pagebreak")
    .Execute Replace:=wdReplaceAll
  End With

' Now that we are sure every PB char has PB style, remove all PB char
  MacroHelpers.zz_clearFind currentRng
  With currentRng.Find
    .Text = "^m"
    .Replacement.Text = ""
    .Execute Replace:=wdReplaceAll
  End With

' Remove multiple PB-styled paragraphs in a row
  MacroHelpers.zz_clearFind currentRng
  With currentRng.Find
    .Text = "^13{2,}"
    .Replacement.Text = "^p"
    .Format = True
    .Style = WT_StyleConfig.MacmillanStyles("pagebreak")
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll
  End With

' Now we should have all PBs are just single paragraph return with PB style.
' So we just need to add the PB character.
  MacroHelpers.zz_clearFind currentRng
  With currentRng.Find
    .Text = "^p"
    .Replacement.Text = "^m^p"
    .Format = True
    .Style = WT_StyleConfig.MacmillanStyles("pagebreak")
    .MatchWildcards = False
    .Execute Replace:=wdReplaceAll
  End With

' Remove any manual breaks preceding section start styles
  Dim varStyle As Variant
  Dim lngSectionStyle As Long
  Dim lngPrevSibling As Long
  Dim colDeleteTheseOnes As Collection
  Set colDeleteTheseOnes = New Collection

  For Each varStyle In WT_StyleConfig.SectionStartStyles
    MacroHelpers.zz_clearFind
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
      .Format = True
      .Style = varStyle
      .Forward = True
      .Wrap = wdFindStop
      .Execute
      
      Do While .Found = True
      ' Check previous paragraph for Page Break
        lngSectionStyle = MacroHelpers.ParaIndex
        If lngSectionStyle > 1 Then
          lngPrevSibling = lngSectionStyle - 1
          If activeDoc.Paragraphs(lngPrevSibling).Range.ParagraphStyle = WT_StyleConfig.MacmillanStyles("pagebreak") Then
          ' Add to collection to delete later (because if we delete it now it will mess up
          ' our paragraph indices. Add Before first so our list is reverse sorted, so we
          ' can delete from the bottom up (again to not mess with paragraph indices)
            colDeleteTheseOnes.Add Item:=lngPrevSibling, Before:=1
          End If
        End If
        .Execute
      Loop
    End With
  Next varStyle
  
' Actually delete these now
  Dim varIndex As Variant
  For Each varIndex In colDeleteTheseOnes
    activeDoc.Paragraphs(varIndex).Range.Delete
  Next varIndex

' Remove any first paragraphs until we get one with text
  Dim lngCount As Long
  Dim rngPara1 As Range
  Dim strKey As String
  
  Do
    lngCount = lngCount + 1
    strKey = "firstParaPB" & lngCount
    Set rngPara1 = activeDoc.Paragraphs.First.Range
    If MacroHelpers.IsNewLine(rngPara1.Text) = True And _
      WT_StyleConfig.IsSectionStartStyle(rngPara1.ParagraphStyle) = False Then
      rngPara1.Delete
    Else
      Exit Do
    End If
  Loop Until lngCount > 20 ' For runaway loops
End Sub




