Attribute VB_Name = "SharedMacros"

' All should be declared as Public for use from other modules

Option Explicit

Public Enum GitBranch
    master = 1
    releases = 2
    develop = 3
End Enum

Public Enum TemplatesList
    updaterTemplates = 1
    toolsTemplates = 2
    stylesTemplates = 3
    installTemplates = 4
    allTemplates = 5
End Enum
Public Function StyleDir() As String
    Dim strFullPath As String
    Dim strMacDocs As String
    Dim strStylesName As String
    
    strStylesName = "MacmillanStyleTemplate"
    
    #If Mac Then
        strMacDocs = MacScript("return (path to documents folder) as string")
        strFullPath = strMacDocs & strStylesName
    #Else
        strFullPath = Environ("APPDATA") & Application.PathSeparator & strStylesName
    #End If
    
'    Debug.Print strFullPath
    StyleDir = strFullPath
    
End Function


Public Function GetTemplatesList(TemplatesYouWant As TemplatesList, Optional PathToRepo As String) As Variant
    ' returns an array of paths to template files in their final installation locations
    ' if you want to use "allTemplates" (i.e., for updating code in templates), must include PathToRepo
    
    Dim strStartupDir As String
    Dim strStyleDir As String
    
    strStartupDir = Application.StartupPath
    strStyleDir = StyleDir()

    Dim strPathsToTemplates() As String
    Dim K As Long
    K = 0
    
    ' get the updater file for these requests
    If TemplatesYouWant = updaterTemplates Or _
        TemplatesYouWant = installTemplates Or _
        TemplatesYouWant = allTemplates Then
        K = K + 1
        ReDim Preserve strPathsToTemplates(1 To K)
        strPathsToTemplates(K) = strStartupDir & Application.PathSeparator & "GtUpdater.dotm"
    End If
    
    ' get the tools file for these requests
    If TemplatesYouWant = toolsTemplates Or _
        TemplatesYouWant = installTemplates Or _
        TemplatesYouWant = allTemplates Then
        K = K + 1
        ReDim Preserve strPathsToTemplates(1 To K)
        strPathsToTemplates(K) = strStyleDir & Application.PathSeparator & "MacmillanGT.dotm"
    End If
    
    ' get the styles files for these requests
    If TemplatesYouWant = stylesTemplates Or _
        TemplatesYouWant = installTemplates Or _
        TemplatesYouWant = allTemplates Then
        K = K + 1
        ReDim Preserve strPathsToTemplates(1 To K)
        strPathsToTemplates(K) = strStyleDir & Application.PathSeparator & "macmillan.dotm"
        
        K = K + 1
        ReDim Preserve strPathsToTemplates(1 To K)
        strPathsToTemplates(K) = strStyleDir & Application.PathSeparator & "macmillan_NoColor.dotm"

        K = K + 1
        ReDim Preserve strPathsToTemplates(1 To K)
        strPathsToTemplates(K) = strStyleDir & Application.PathSeparator & "macmillan_CoverCopy.dotm"
    End If
    
    ' also get the installer file
    If TemplatesYouWant = allTemplates And PathToRepo <> vbNullString Then
        K = K + 1
        ReDim Preserve strPathsToTemplates(1 To K)
        strPathsToTemplates(K) = PathToRepo & Application.PathSeparator & "MacmillanTemplateInstaller" _
            & Application.PathSeparator & "MacmillanTemplateInstaller.docm"
        
        ' Could also add paths to open _BETA and _DEVELOP installer files?
    End If
    
    ' DEBUGGING: check tha list!
'    Dim H As Long
'    For H = LBound(strPathsToTemplates) To (UBound(strPathsToTemplates))
'        Debug.Print H & ": " & strPathsToTemplates(H)
'    Next H
    
    
    GetTemplatesList = strPathsToTemplates
    
End Function


Public Function IsItThere(Path)
' Check if file or directory exists on PC or Mac
    
    'Debug.Print Path
    
    'Remove trailing path separator from dir if it's there
    If Right(Path, 1) = Application.PathSeparator Then
        Path = Left(Path, Len(Path) - 1)
    End If
    
    Dim CheckDir As String
    On Error GoTo ErrHandler            ' Because Dir(Path) throws an error on Mac if not existant
        
    CheckDir = Dir(Path, vbDirectory)
    
    If CheckDir = vbNullString Then
        IsItThere = False
    Else
        IsItThere = True
    End If
    
    On Error GoTo 0
    
Exit Function

ErrHandler:
    If Err.Number = 68 Then     ' "Device unavailable"
        IsItThere = False
    Else
        Debug.Print "IsItThere Error " & Err.Number & ": " & Err.Description
    End If
End Function

' ===== DownloadFromConfluence ================================================
' Actually now it downloads from Github but don't want to mess with things, we're
' going to be totally refacroting soon.
Public Function DownloadFromConfluence(FinalDir As String, LogFile As String, FileName As String, _
    Optional DownloadSource As GitBranch = master) As Boolean
'FinalDir is directory w/o file name

    Dim logString As String
    Dim strMacTmpDir As String
    Dim strTmpPath As String
    Dim strBashTmp As String
    Dim strFinalPath As String
    Dim strErrMsg As String
    Dim myURL As String
    Dim strBranch As String
    Dim strDownloadRepo As String
    Dim strBaseUrl As String
    Dim strSubfolder As String
    
    strFinalPath = FinalDir & Application.PathSeparator & FileName

'Get URL to download from. Hard coded for now since will be replaced with config refactor
    ' Base URL everything is available from
    strBaseUrl = "https://raw.githubusercontent.com/macmillanpublishers/"
    
    ' Branch to download from
    Select Case DownloadSource
      Case develop
        strBranch = "develop/"
      Case master
        strBranch = "master/"
      Case releases
        strBranch = "releases/"
    End Select
    
    ' Determine repo and file path from file name. Will be handled better in config.
    If InStr(1, FileName, "gt", vbTextCompare) Then
      strDownloadRepo = "Word-template/"
      strSubfolder = Left(FileName, InStr(FileName, ".") - 1) & "/"
    Else
      strDownloadRepo = "Word-template_assets/"
      strSubfolder = vbNullString
    End If
    
    ' put it all together
    myURL = strBaseUrl & strDownloadRepo & strBranch & strSubfolder & FileName
    Debug.Print "Attempting to download: " & myURL
    
    'Get temp dir based on OS, then download file.
    #If Mac Then
        'set tmp dir
        strMacTmpDir = MacScript("path to temporary items as string")
        strTmpPath = strMacTmpDir & FileName
        'Debug.Print strTmpPath
        strBashTmp = Replace(Right(strTmpPath, Len(strTmpPath) - (InStr(strTmpPath, ":") - 1)), ":", "/")
        'Debug.Print strBashTmp
        
        'check for network
        If ShellAndWaitMac("ping -o google.com &> /dev/null ; echo $?") <> 0 Then   'can't connect to internet
            logString = Now & " -- Tried update; unable to connect to network."
            LogInformation LogFile, logString
            strErrMsg = "There was an error trying to download the Macmillan template." & vbNewLine & vbNewLine & _
                        "Please check your internet connection or contact workflows@macmillan.com for help."
            MsgBox strErrMsg, vbCritical, "Error 1: Connection error (" & FileName & ")"
            DownloadFromConfluence = False
            Exit Function
        Else 'internet is working, download file
            'Make sure file is there
            Dim httpStatus As Long
            httpStatus = ShellAndWaitMac("curl -s -o /dev/null -w '%{http_code}' " & myURL)
            
            If httpStatus = 200 Then                    ' File is there
                'Now delete file if already there, then download new file
                ShellAndWaitMac ("rm -f " & strBashTmp & " ; curl -o " & strBashTmp & " " & myURL)
            ElseIf httpStatus = 404 Then            ' 404 = page not found
                logString = Now & " -- 404 File not found. Cannot download file."
                LogInformation LogFile, logString
                strErrMsg = "It looks like that file isn't available for download." & vbNewLine & vbNewLine & _
                    "Please contact workflows@macmillan.com for help."
                MsgBox strErrMsg, vbCritical, "Error 7: File not found (" & FileName & ")"
                DownloadFromConfluence = False
                Exit Function
            Else
                logString = Now & " -- Http status is " & httpStatus & ". Cannot download file."
                LogInformation LogFile, logString
                strErrMsg = "There was an error trying to download the Macmillan templates." & vbNewLine & vbNewLine & _
                    "Please check your internet connection or contact workflows@macmillan.com for help."
                MsgBox strErrMsg, vbCritical, "Error 2: Http status " & httpStatus & " (" & FileName & ")"
                DownloadFromConfluence = False
                Exit Function
            End If

        End If
    #Else
        'set tmp dir
        strTmpPath = Environ("TEMP") & Application.PathSeparator & FileName 'Environ gives temp dir for Mac too? NOPE
        
        'Check if file is already in tmp dir, delete if yes
        If IsItThere(strTmpPath) = True Then
            Kill strTmpPath
        End If
        
        'try to download the file from Public Confluence page
        Dim WinHttpReq As Object
        Dim oStream As Object
        
        'Attempt to download file
        On Error Resume Next
            Set WinHttpReq = CreateObject("MSXML2.XMLHTTP.3.0")
            WinHttpReq.Open "GET", myURL, False
            WinHttpReq.Send
    
                ' Exit sub if error in connecting to website
                If Err.Number <> 0 Then 'HTTP request is not OK
                    'Debug.Print WinHttpReq.Status
                    logString = Now & " -- could not connect to Confluence site: Error " & Err.Number
                    LogInformation LogFile, logString
                    strErrMsg = "There was an error trying to download the Macmillan template." & vbNewLine & vbNewLine & _
                        "Please check your internet connection or contact workflows@macmillan.com for help."
                    MsgBox strErrMsg, vbCritical, "Error 1: Connection error (" & FileName & ")"
                    DownloadFromConfluence = False
                    On Error GoTo 0
                    Exit Function
                End If
        On Error GoTo 0
        
        'Debug.Print "Http status for " & FileName & ": " & WinHttpReq.Status
        If WinHttpReq.Status = 200 Then  ' 200 = HTTP request is OK
        
            'if connection OK, download file to temp dir
            myURL = WinHttpReq.responseBody
            Set oStream = CreateObject("ADODB.Stream")
            oStream.Open
            oStream.Type = 1
            oStream.Write WinHttpReq.responseBody
            oStream.SaveToFile strTmpPath, 2 ' 1 = no overwrite, 2 = overwrite
            oStream.Close
            Set oStream = Nothing
            Set WinHttpReq = Nothing
        ElseIf WinHttpReq.Status = 404 Then ' 404 = file not found
            logString = Now & " -- 404 File not found. Cannot download file."
            LogInformation LogFile, logString
            strErrMsg = "It looks like that file isn't available for download." & vbNewLine & vbNewLine & _
                "Please contact workflows@macmillan.com for help."
            MsgBox strErrMsg, vbCritical, "Error 7: File not found (" & FileName & ")"
            DownloadFromConfluence = False
            Exit Function
        Else
            logString = Now & " -- Http status is " & WinHttpReq.Status & ". Cannot download file."
            LogInformation LogFile, logString
            strErrMsg = "There was an error trying to download the Macmillan templates." & vbNewLine & vbNewLine & _
                "Please check your internet connection or contact workflows@macmillan.com for help."
            MsgBox strErrMsg, vbCritical, "Error 2: Http status " & WinHttpReq.Status & " (" & FileName & ")"
            DownloadFromConfluence = False
            Exit Function
        End If
    #End If
        
    'Error if download was not successful
    If IsItThere(strTmpPath) = False Then
        logString = Now & " -- " & FileName & " file download to Temp was not successful."
        LogInformation LogFile, logString
        strErrMsg = "There was an error downloading the Macmillan template." & vbNewLine & _
            "Please contact workflows@macmillan.com for assitance."
        MsgBox strErrMsg, vbCritical, "Error 3: Download failed (" & FileName & ")"
        DownloadFromConfluence = False
        On Error GoTo 0
        Exit Function
    Else
        logString = Now & " -- " & FileName & " file download to Temp was successful."
        LogInformation LogFile, logString
    End If


    
    'If file exists already, log it and delete it
    If IsItThere(strFinalPath) = True Then

        logString = Now & " -- Previous version file in final directory."
        LogInformation LogFile, logString
        
        ' get file extension
        Dim strExt As String
        strExt = Right(strFinalPath, InStrRev(StrReverse(strFinalPath), "."))
        
        ' can't delete template if it's installed as an add-in
        If InStr(strExt, "dot") > 0 Then
            On Error Resume Next        'Error = add-in not available, don't need to uninstall
                AddIns(strFinalPath).Installed = False
            On Error GoTo 0
        End If
  
        ' Test if dir is read only
        If IsReadOnly(FinalDir) = True Then ' Dir is read only
            logString = Now & " -- old " & FileName & " file is read only, can't delete/replace. " _
                & "Alerting user."
            LogInformation LogFile, logString
            strErrMsg = "The installer doesn't have permission. Please conatct workflows" & _
                "@macmillan.com for help."
            MsgBox strErrMsg, vbCritical, "Error 8: Permission denied (" & FileName & ")"
            DownloadFromConfluence = False
            On Error GoTo 0
            Exit Function
        Else
            On Error Resume Next
                Kill strFinalPath
                
                If Err.Number = 70 Then         'File is open and can't be replaced
                    logString = Now & " -- old " & FileName & " file is open, can't delete/replace. Alerting user."
                    LogInformation LogFile, logString
                    strErrMsg = "Please close all other Word documents and try again."
                    MsgBox strErrMsg, vbCritical, "Error 4: Previous version removal failed (" & FileName & ")"
                    DownloadFromConfluence = False
                    On Error GoTo 0
                    Exit Function
                End If
            On Error GoTo 0
        End If
    Else
        logString = Now & " -- No previous version file in final directory."
        LogInformation LogFile, logString
    End If
        
    'If delete was successful, move downloaded file to final directory
    If IsItThere(strFinalPath) = False Then
        logString = Now & " -- Final directory clear of " & FileName & " file."
        LogInformation LogFile, logString
        
        ' move template to final directory
        Name strTmpPath As strFinalPath
        
        'Mac won't load macros from a template downloaded from the internet to Startup.
        'Need to send these commands for it to work, see Confluence
        ' Do NOT use open/save as option, this removes customUI which creates Mac Tools toolbar later
        #If Mac Then
            If InStr(1, FileName, ".dotm") Then
            Dim strCommand As String
            strCommand = "do shell script " & Chr(34) & "xattr -wx com.apple.FinderInfo \" & Chr(34) & _
                "57 58 54 4D 4D 53 57 44 00 10 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00\" & _
                Chr(34) & Chr(32) & Chr(34) & " & quoted form of POSIX path of " & Chr(34) & strFinalPath & Chr(34)
                'Debug.Print strCommand
                MacScript (strCommand)
            End If
        #End If
    
    Else
        logString = Now & " -- old " & FileName & " file not cleared from Final directory."
        LogInformation LogFile, logString
        strErrMsg = "There was an error installing the Macmillan template." & vbNewLine & _
            "Please close all other Word documents and try again, or contact workflows@macmillan.com."
        MsgBox strErrMsg, vbCritical, "Error 5: Previous version uninstall failed (" & FileName & ")"
        DownloadFromConfluence = False
        On Error GoTo 0
        Exit Function
    End If
    
    'If move was successful, yay! Else, :(
    If IsItThere(strFinalPath) = True Then
        logString = Now & " -- " & FileName & " file successfully saved to final directory."
        LogInformation LogFile, logString
    Else
        logString = Now & " -- " & FileName & " file not saved to final directory."
        LogInformation LogFile, logString
        strErrMsg = "There was an error installing the Macmillan template." & vbNewLine & vbNewLine & _
            "Please cotact workflows@macmillan.com for assistance."
        MsgBox strErrMsg, vbCritical, "Error 6: Installation failed (" & FileName & ")"
        DownloadFromConfluence = False
        On Error GoTo 0
        Exit Function
    End If
    
    'Cleanup: Get rid of temp file if downloaded correctly
    If IsItThere(strTmpPath) = True Then
        Kill strTmpPath
    End If
    
    ' Disable Startup add-ins so they don't launch right away and mess of the code that's running
    If InStr(1, LCase(strFinalPath), LCase("startup"), vbTextCompare) > 0 Then         'LCase because "startup" was staying in all caps for some reason, UCase wasn't working
        On Error Resume Next                                        'Error = add-in not available, don't need to uninstall
            AddIns(strFinalPath).Installed = False
        On Error GoTo 0
    End If
    
    DownloadFromConfluence = True

End Function
 
Public Function ShellAndWaitMac(Cmd As String) As String

    Dim result As String
    Dim scriptCmd As String ' Macscript command
    
    scriptCmd = "do shell script """ & Cmd & """"
    result = MacScript(scriptCmd) ' result contains stdout, should you care
    'Debug.Print result
    ShellAndWaitMac = result

End Function

Public Sub LogInformation(LogFile As String, LogMessage As String)

Dim FileNum As Integer
    FileNum = FreeFile ' next file number
    Open LogFile For Append As #FileNum ' creates the file if it doesn't exist
    Print #FileNum, LogMessage ' write information at the end of the text file
    Close #FileNum ' close the file
End Sub

Public Sub OverwriteTextFile(TextFile As String, NewText As String)
' TextFile should be full path
    
    Dim FileNum As Integer
    
    If IsItThere(TextFile) = True Then
        FileNum = FreeFile ' next file number
        Open TextFile For Output Access Write As #FileNum
        Print #FileNum, NewText ' overwrite information in the text of the file
        Close #FileNum ' close the file
    End If

End Sub

Public Function CreateLogFileInfo(ByRef FileName As String) As Variant
' Creates the style dir, log dir, and log file name variables for use in other subs.
' File name should not contain periods other than before file type

    Dim strLogFile As String
    Dim strMacDocs As String
    Dim strStyle As String
    Dim strLogFolder As String
    Dim strLogPath As String
    
    'Create logfile name
    strLogFile = Left(FileName, InStrRev(FileName, ".") - 1)
    strLogFile = strLogFile & "_updates.log"
    strStyle = StyleDir()
    strLogFolder = strStyle & Application.PathSeparator & "log"
    strLogPath = strLogFolder & Application.PathSeparator & strLogFile

    'Debug.Print strLogPath

    Dim arrFinalDirs() As Variant
    ReDim arrFinalDirs(1 To 3)
    
    arrFinalDirs(1) = strStyle
    arrFinalDirs(2) = strLogFolder
    arrFinalDirs(3) = strLogPath
    
    CreateLogFileInfo = arrFinalDirs

End Function


'Public Function NotesExist(StoryType As WdStoryType) As Boolean
'    On Error GoTo ErrHandler
'    Dim myRange As Range
'    Set myRange = ActiveDocument.StoryRanges(StoryType)
'    'If can set as myRange, then exists
'    NotesExist = True
'    On Error GoTo 0
'    Exit Function
'ErrHandler:
'    If Err.Number = 5941 Then   '"Member of the collection does not exist"
'        NotesExist = False
'    End If
'End Function

Public Sub zz_clearFind()

    Dim clearRng As Range
    Set clearRng = ActiveDocument.Words.First

    With clearRng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute
    End With
    
End Sub

Public Function StoryArray() As Variant
    '------------check for endnotes and footnotes--------------------------
    Dim strStories() As Variant
    
    ReDim strStories(1 To 1)
    strStories(1) = wdMainTextStory
    
    If ActiveDocument.Endnotes.Count > 0 Then
        ReDim Preserve strStories(1 To (UBound(strStories()) + 1))
        strStories(UBound(strStories())) = wdEndnotesStory
    End If
    
    If ActiveDocument.Footnotes.Count > 0 Then
        ReDim Preserve strStories(1 To (UBound(strStories()) + 1))
        strStories(UBound(strStories())) = wdFootnotesStory
    End If
    
    StoryArray = strStories
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
            lngLines = ActiveDocument.ComputeStatistics(wdStatisticLines)
            
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

Function CheckSave()
    ' Prompts user to save document before running the macro. If they click "Cancel" then CheckSave returns true and
    ' you should exit your macro. also checks if document protection is on.
    
    Dim mainDoc As Document
    Set mainDoc = ActiveDocument
    Dim iReply As Integer
    
    '-----make sure document is saved
    Dim docSaved As Boolean                                                                                                 'v. 3.1 update
    docSaved = mainDoc.Saved
    
    If docSaved = False Then
        iReply = MsgBox("Your document '" & mainDoc & "' contains unsaved changes." & vbNewLine & vbNewLine & _
            "Click OK to save your document and run the macro." & vbNewLine & vbNewLine & "Click 'Cancel' to exit.", _
                vbOKCancel, "Error 1")
        If iReply = vbOK Then
            CheckSave = False
            mainDoc.Save
        Else
            CheckSave = True
            Exit Function
        End If
    End If
    
    '-----test protection
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "Uh oh ... protection is enabled on document '" & mainDoc & "'." & vbNewLine & _
            "Please unprotect the document and run the macro again." & vbNewLine & vbNewLine & _
            "TIP: If you don't know the protection password, try pasting contents of this file into " & _
            "a new file, and run the macro on that.", , "Error 2"
        CheckSave = True
        Exit Function
    Else
        CheckSave = False
    End If

End Function


Public Function IsStyleInDoc(StyleName As Variant) As Boolean
  On Error GoTo IsStyleInDocError
  Dim blnResult As Boolean: blnResult = True
  Dim TestStyle As Style
  
' Try to access this style. If not present in doc, will error
  Set TestStyle = ActiveDocument.Styles.Item(StyleName)
  IsStyleInDoc = blnResult
  Exit Function
  
IsStyleInDocError:
' 5941 = "The requested member of the collection does not exist."
' Have to test here, ErrorChecker tries to create style if missing
  If Err.Number = 5941 Then
    blnResult = False
    Resume Next
  Else
    Exit Function
  End If
' Otherwise, usual error stuff:
  
End Function

Function IsArrayEmpty(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' By Chip Pearson, http://www.cpearson.com/excel/vbaarrays.htm
'
' IsArrayEmpty
' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is really the reverse of IsArrayAllocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim LB As Long
    Dim UB As Long
    
    Err.Clear
    On Error Resume Next
    If IsArray(Arr) = False Then
        ' we weren't passed an array, return True
        IsArrayEmpty = True
        Exit Function
    End If
    
    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    UB = UBound(Arr, 1)
    If (Err.Number <> 0) Then
        IsArrayEmpty = True
    Else
        ''''''''''''''''''''''''''''''''''''''''''
        ' On rare occassion, under circumstances I
        ' cannot reliably replictate, Err.Number
        ' will be 0 for an unallocated, empty array.
        ' On these occassions, LBound is 0 and
        ' UBoung is -1.
        ' To accomodate the weird behavior, test to
        ' see if LB > UB. If so, the array is not
        ' allocated.
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        LB = LBound(Arr)
        If LB > UB Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If

End Function


Function GetText(StyleName As String) As String
    Dim fString As String
    Dim fCount As Integer
    
    Application.ScreenUpdating = False
    
    fCount = 0
    
    'Move selection to start of document
    Selection.HomeKey Unit:=wdStory
    
    On Error GoTo ErrHandler
    
        Selection.Find.ClearFormatting
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .Style = ActiveDocument.Styles(StyleName)
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
        fString = fString & Selection.Text & vbNewLine
        
        'If the next character is a paragraph return, add that to the selection
        'Otherwise the next Find will just select the same text with the paragraph return
        If InStr(StyleName, "span") = 0 Then        'Don't select terminal para mark if char style, sends into an infinite loop
            Selection.MoveEndWhile Cset:=Chr(13), Count:=1
        End If
    Loop
        
    If fCount = 0 Then
        GetText = ""
    Else
        GetText = fString
    End If
    
    Application.ScreenUpdating = True
    
    Exit Function
    
ErrHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then   ' The style is not present in the document
        GetText = ""
    End If
        
End Function

Function LoadCSVtoArray(Path As String, RemoveHeaderRow As Boolean, RemoveHeaderCol As Boolean) As Variant

'------Load CSV into 2d array, NOTE!!: base 0---------
' But also note that this now removes the header row and column too
    Dim fnum As Integer
    Dim whole_file As String
    Dim lines As Variant
    Dim one_line As Variant
    Dim num_rows As Long
    Dim num_cols As Long
    Dim the_array() As Variant
    Dim R As Long
    Dim C As Long
    
        If IsItThere(Path) = False Then
            MsgBox "There was a problem with your Castoff.", vbCritical, "Error: CSV not available"
            Exit Function
        End If
        'Debug.Print Path
        
        ' Do we need to remove a header row?
        Dim lngHeaderRow As Long
        If RemoveHeaderRow = True Then
            lngHeaderRow = 1
        Else
            lngHeaderRow = 0
        End If
        
        ' Do we need to remove a header column?
        Dim lngHeaderCol As Long
        If RemoveHeaderCol = True Then
            lngHeaderCol = 1
        Else
            lngHeaderCol = 0
        End If
        
        ' Load the csv file.
        fnum = FreeFile
        Open Path For Input As fnum
        whole_file = Input$(LOF(fnum), #fnum)
        Close fnum

        ' Break the file into lines (trying to capture whichever line break is used)
        If InStr(1, whole_file, vbCrLf) <> 0 Then
            lines = Split(whole_file, vbCrLf)
        ElseIf InStr(1, whole_file, vbCr) <> 0 Then
            lines = Split(whole_file, vbCr)
        ElseIf InStr(1, whole_file, vbLf) <> 0 Then
            lines = Split(whole_file, vbLf)
        Else
            MsgBox "There was an error with your castoff.", vbCritical, "Error parsing CSV file"
        End If

        ' Dimension the array.
        num_rows = UBound(lines)
        one_line = Split(lines(0), ",")
        num_cols = UBound(one_line)
        ReDim the_array(num_rows - lngHeaderRow, num_cols - lngHeaderCol) ' -1 if we are not using header row/col
        
        ' Copy the data into the array.
        For R = lngHeaderRow To num_rows           ' start at 1 (not 0) if we are not using the header row
            If Len(lines(R)) > 0 Then
                one_line = Split(lines(R), ",")
                For C = lngHeaderCol To num_cols   ' start at 1 (not 0) if we are not using the header column
                    'Debug.Print one_line(c)
                    the_array((R - lngHeaderRow), (C - lngHeaderCol)) = one_line(C)   ' -1 because if are not using header row/column from CSV
                Next C
            End If
        Next R
    
        ' Prove we have the data loaded.
'         Debug.Print LBound(the_array)
'         Debug.Print UBound(the_array)
'         For R = 0 To (num_rows - 1)          ' -1 again if we removed the header row
'             For c = 0 To num_cols      ' -1 again if we removed the header column
'                 Debug.Print the_array(R, c) & " | ";
'             Next c
'             Debug.Print
'         Next R
'         Debug.Print "======="
    
    LoadCSVtoArray = the_array
 
End Function

Private Function FixTrackChanges() As Boolean
    Dim N As Long
    Dim oComments As Comments
    Set oComments = ActiveDocument.Comments
    
    Application.ScreenUpdating = False
    
    FixTrackChanges = True
    
    Application.DisplayAlerts = False
    
    'See if there are tracked changes or comments in document
    On Error Resume Next
    Selection.HomeKey Unit:=wdStory   'start search at beginning of doc
    WordBasic.NextChangeOrComment       'search for a tracked change or comment. error if none are found.
    
    'If there are changes, ask user if they want macro to accept changes or cancel
    If Err = 0 Then
        If MsgBox("Bookmaker doesn't like comments or tracked changes, but it appears that you have some in your document." _
            & vbCr & vbCr & "Click OK to ACCEPT ALL CHANGES and DELETE ALL COMMENTS right now and continue with the Bookmaker Requirements Check." _
            & vbCr & vbCr & "Click CANCEL to stop the Bookmaker Requirements Check and deal with the tracked changes and comments on your own.", _
            273, "Are those tracked changes I see?") = vbCancel Then           '273 = vbOkCancel(1) + vbCritical(16) + vbDefaultButton2(256)
                FixTrackChanges = False
                Exit Function
        Else 'User clicked OK, so accept all tracked changes and delete all comments
            ActiveDocument.AcceptAllRevisions
            For N = oComments.Count To 1 Step -1
                oComments(N).Delete
            Next N
            Set oComments = Nothing
        End If
    End If
    
    On Error GoTo 0
    Application.DisplayAlerts = True
    
End Function

Function IsReadOnly(Path As String) As Boolean
    ' Tests if the file or directory is read-only
    
    #If Mac Then
        Dim strScript As String
        Dim blnWritable As Boolean
        
        strScript = _
            "set p to POSIX path of " & Chr(34) & Path & Chr(34) & Chr(13) & _
            "try" & Chr(13) & _
            vbTab & "do shell script " & Chr(34) & "test -w \" & Chr(34) & "$(dirname " & Chr(34) & _
                " & quoted form of p & " & Chr(34) & ")\" & Chr(34) & Chr(34) & Chr(13) & _
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
    #Else
        If (GetAttr(Path) And vbReadOnly) <> 0 Then
            IsReadOnly = True
        Else
            IsReadOnly = False
        End If
    #End If
    
End Function


Public Function ReadTextFile(Path As String, Optional FirstLineOnly As Boolean = True) As String
' load string from text file

    Dim fnum As Long
    Dim strTextWeWant As String
    
    fnum = FreeFile()
    Open Path For Input As fnum
    
    If FirstLineOnly = False Then
        strTextWeWant = Input$(LOF(fnum), #fnum)
    Else
        Line Input #fnum, strTextWeWant
    End If
    
    Close fnum
    
    ReadTextFile = strTextWeWant
End Function


Function HiddenTextSucks(StoryType As WdStoryType) As Boolean                                             'v. 3.1 patch : redid this whole thing as an array, addedsmart quotes, wrap toggle var
'    Debug.Print StoryType
    Dim activeRng As Range
    Set activeRng = ActiveDocument.StoryRanges(StoryType)
    ' No, really, it does. Why is that even an option?
    ' Seriously, this just deletes all hidden text, based on the
    ' assumption that if it's hidden, you don't want it.
    ' returns a Boolean in case we want to notify user at some point
    
    HiddenTextSucks = False
    
    ' If Hidden text isn't shown, it won't be deleted, which
    ' defeats the purpose of doing this at all.
    Dim blnCurrentHiddenView As Boolean
    blnCurrentHiddenView = ActiveDocument.ActiveWindow.View.ShowAll
    ActiveDocument.ActiveWindow.View.ShowAll = True

    
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
    ActiveDocument.ActiveWindow.View.ShowAll = blnCurrentHiddenView
    
End Function

Sub StyleAllHyperlinks(StoriesInUse As Variant)
    ' StoriesInUse is an array of wdStoryTypes in use
    ' Clears active links and adds macmillan URL char styles
    ' to any proper URLs.
    ' Breaking up into sections because AutoFormat does not apply hyperlinks to FN/EN stories.
    ' Also if you AutoFormat a second time it undoes all of the formatting already applied to hyperlinks
    
    Dim S As Long
    
    Call zz_clearFind
    
    For S = 1 To UBound(StoriesInUse)
        'Styles hyperlinks, must be performed after PreserveWhiteSpaceinBrkStylesA
        Call SharedMacros.StyleHyperlinksA(StoryType:=(StoriesInUse(S)))
    Next S
    
    Call SharedMacros.AutoFormatHyperlinks
    
    For S = 1 To UBound(StoriesInUse)
        Call SharedMacros.StyleHyperlinksB(StoryType:=(StoriesInUse(S)))
    Next S
    
End Sub

Private Sub StyleHyperlinksA(StoryType As WdStoryType)
    ' PRIVATE, if you want to style hyperlinks from another module,
    ' call StyleAllHyperlinks sub above.
    ' added by Erica 2014-10-07, v. 3.4
    ' removes all live hyperlinks but leaves hyperlink text intact
    ' then styles all URLs as "span hyperlink (url)" style
    ' -----------------------------------------
    ' this first bit removes all live hyperlinks from document
    ' we want to remove these from urls AND text; will add back to just urls later
    Dim activeRng As Range
    Set activeRng = ActiveDocument.StoryRanges(StoryType)
    ' remove all embedded hyperlinks regardless of character style
    With activeRng
        While .Hyperlinks.Count > 0
            .Hyperlinks(1).Delete
        Wend
    End With
    '------------------------------------------
    'removes all hyperlink styles
    Dim HyperlinkStyleArray(3) As String
    Dim P As Long
    
On Error GoTo LinksErrorHandler:
    
    HyperlinkStyleArray(1) = "Hyperlink"        'built-in style applied automatically to links
    HyperlinkStyleArray(2) = "FollowedHyperlink"    'built-in style applied automatically
    HyperlinkStyleArray(3) = "span hyperlink (url)" 'Macmillan template style for links
    
    For P = 1 To UBound(HyperlinkStyleArray())
        With activeRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Style = HyperlinkStyleArray(P)
            .Replacement.Style = ActiveDocument.Styles("Default Paragraph Font")
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
    
On Error GoTo 0
    
    Exit Sub
    
LinksErrorHandler:
        '5834 means item does not exist
        '5941 means style not present in collection
        If Err.Number = 5834 Or Err.Number = 5941 Then
            
            'If style is not present, add style
            Dim myStyle As Style
            Set myStyle = ActiveDocument.Styles.Add(Name:="span hyperlink (url)", Type:=wdStyleTypeCharacter)
            Resume
'            ' Used to add highlight color, but actually if style is missing, it's
'            ' probably a MS w/o Macmillan's styles and the highlight will be annoying.
'            'If missing style was Macmillan built-in style, add character highlighting
'            If myStyle = "span hyperlink (url)" Then
'                ActiveDocument.Styles("span hyperlink (url)").Font.Shading.BackgroundPatternColor = wdColorPaleBlue
'            End If
        Else
            MsgBox "Error " & Err.Number & ": " & Err.Description
            On Error GoTo 0
            Exit Sub
        End If

End Sub

Private Sub AutoFormatHyperlinks()
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
    Dim f10 As Boolean
      
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
        ' Perform AutoFormat
        ActiveDocument.Content.AutoFormat
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
    End With
    
    'This bit autoformats hyperlinks in endnotes and footnotes
    ' from http://www.vbaexpress.com/forum/showthread.php?52466-applying-hyperlink-styles-in-footnotes-and-endnotes
    Dim oDoc As Document
    Dim oTemp As Document
    Dim oNote As Range
    Dim oRng As Range
    
    'oDoc.Save      ' Already saved active doc?
    Set oDoc = ActiveDocument
    Set oTemp = Documents.Add(Template:=oDoc.FullName, Visible:=False)
    
    If oDoc.Footnotes.Count >= 1 Then
        Dim oFN As Footnote
        For Each oFN In oDoc.Footnotes
            Set oNote = oFN.Range
            Set oRng = oTemp.Range
            oRng.FormattedText = oNote.FormattedText
            'oRng.Style = "Footnote Text"
            Options.AutoFormatReplaceHyperlinks = True
            oRng.AutoFormat
            oRng.End = oRng.End - 1
            oNote.FormattedText = oRng.FormattedText
        Next oFN
        Set oFN = Nothing
    End If
    
    If oDoc.Endnotes.Count >= 1 Then
        Dim oEN As Endnote
        For Each oEN In oDoc.Endnotes
            Set oNote = oEN.Range
            Set oRng = oTemp.Range
            oRng.FormattedText = oNote.FormattedText
            'oRng.Style = "Endnote Text"
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
    
End Sub

Private Sub StyleHyperlinksB(StoryType As WdStoryType)
    ' PRIVATE, if you want to style hyperlinks from another module,
    ' call StyleAllHyperlinks sub above.
    '--------------------------------------------------
    ' apply macmillan URL style to hyperlinks we just tagged in Autoformat
    Dim activeRng As Range
    Set activeRng = ActiveDocument.StoryRanges(StoryType)
    With activeRng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Style = "Hyperlink"
        .Replacement.Style = ActiveDocument.Styles("span hyperlink (url)")
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
    
End Sub

