Attribute VB_Name = "FileInstaller"
Option Explicit
Option Base 1

Sub Installer(Installer As Boolean, TemplateName As String, ByRef FileName() As String, ByRef FinalDir() As String)
'created by Erica Warren - erica.warren@macmillan.com

'======== PURPOSE ===================================
'Downloads and installs an array of template files & logs the downloads

'======== USE =======================================
'This is Part 2 of 2. Must be called from a sub in another module that declares file names and locations.
'The template file needs to be uploaded as an attachment to https://confluence.macmillan.com/display/PBL/Test
'If this is an installer file, The Part 1 code needs to reside in the ThisDocument module as a sub called
'Documents_Open in a .docm file so that it will launch when users open the file.

'"Installer" argument = True if this is for a standalone installtion file.
'"Installer" argument = False is this is part of a daily check of the current file and only updates if out of date.
    
    '' --------------- Check that variables were passed correctly -------------------------
    'Dim x As Long
    'For x = LBound(FileName()) To UBound(FileName())
    '    Debug.Print & " " & FileName(x) & " " & FinalDir(x) & vbNewLine
    'Next x
    
    '' --------------- Set up variable names ----------------------------------------------
    '' Create style directory and logfile names
    Dim a As Long
    Dim strLogFile() As String
    ReDim strLogFile(LBound(FileName()) To UBound(FileName()))
    Dim strFullLogPath() As String
    ReDim strFullLogPath(LBound(FileName()) To UBound(FileName()))
    Dim strStyleDir As String
    Dim strLogDir As String

    For a = LBound(FileName()) To UBound(FileName())
        ' Remove trailing path separator from dir if it's there, so we know we won't duplicate below
        If Right(FinalDir(a), 1) = Application.PathSeparator Then
            FinalDir(a) = Left(FinalDir(a), Len(FinalDir(a)) - 1)
        End If
        
        ' Create full path to template file
        Dim strTemplatePath() As String
        ReDim strTemplatePath(LBound(FileName()) To UBound(FileName()))
        strTemplatePath(a) = FinalDir(a) & Application.PathSeparator & FileName(a)
        Debug.Print "Full template path: " & strTemplatePath(a)
        
        'Create logfile name
        strLogFile(a) = Left(FileName(a), InStrRev(FileName(a), ".do") - 1)
        strLogFile(a) = strLogFile(a) & "_updates.log"
        
        'Create directory names based on OS
        #If Mac Then
            Dim strUser As String
            strUser = MacScript("tell application " & Chr(34) & "System Events" & Chr(34) & Chr(13) & _
                "return (name of current user)" & Chr(13) & "end tell")
            strStyleDir = "Macintosh HD:Users:" & strUser & ":Documents:MacmillanStyleTemplate"
            strLogDir = strStyleDir & Application.PathSeparator & "log"
            strFullLogPath(a) = strLogDir & Application.PathSeparator & strLogFile(a)
        #Else
            strStyleDir = Environ("ProgramData") & "\MacmillanStyleTemplate"
            strLogDir = strStyleDir & Application.PathSeparator & "log"
            strFullLogPath(a) = strLogDir & Application.PathSeparator & strLogFile(a)
        #End If
        'Debug.Print strFullLogPath(a)
    Next a
    
    ' ------------- Check if we need to do an installation ---------------------------
    ' Check if template exists
    Dim installCheck() As Boolean
    ReDim installCheck(LBound(FileName()) To UBound(FileName()))
    Dim blnTemplateExists() As Boolean
    ReDim blnTemplateExists(LBound(FileName()) To UBound(FileName()))
    Dim blnLogUpToDate() As Boolean
    ReDim blnLogUpToDate(LBound(FileName()) To UBound(FileName()))
    Dim b As Long
    
    For b = LBound(FileName()) To UBound(FileName())
    
        'Check if log dir/file exists, create if it doesn't, check last mod date if it does
        ' If last mod date less than 1 day ago, CheckLog = True
        blnLogUpToDate(b) = CheckLog(strStyleDir, strLogDir, strFullLogPath(b))
        Debug.Print FileName(b) & " log exists and was checked today: " & blnLogUpToDate(b)
        
        ' Check if template exists, if not create any missing directories
        blnTemplateExists(b) = IsTemplateThere(FinalDir(b), FileName(b), strFullLogPath(b))
        Debug.Print FileName(b) & " exists: " & blnTemplateExists(b)
                
        If Installer = False Then 'Because if it's an installer, we just want to install the file
            If blnLogUpToDate(b) = True And blnTemplateExists(b) = True Then ' already checked today, already exists
                installCheck(b) = False
            ElseIf blnLogUpToDate(b) = False And blnTemplateExists(b) = True Then 'Log is new or not checked today, already exists
                'check version number
                installCheck(b) = NeedUpdate(FinalDir(b), FileName(b), strFullLogPath(b))
            Else ' blnTemplateExists = False, just download new template
                 installCheck(b) = True
            End If
        Else
            installCheck(b) = True
        End If
        
    Next b

    ' ---------------- Create new array of template files we need to install -----------------
    Dim strInstallMe() As String
    Dim c As Long
    Dim x As Long
    
    x = 0
    
    For c = LBound(FileName()) To UBound(FileName())
        If installCheck(c) = True Then
            x = x + 1
            ReDim Preserve strInstallMe(1 To x)
            strInstallMe(x) = strTemplatePath(c)
        End If
    Next c
    
    ' ---------------- Check if new array is allocated -----------------------------------
    If IsArrayEmpty(strInstallMe()) = True Then
        If Installer = True Then
            'Application.Quit (wdDoNotSaveChanges)
        Else
            Exit Sub
        End If
    Else ' There are values in the array and we need to install them
    
        ' Alert user that installation is happening
        Dim strWelcome As String
    
        strWelcome = "Welcome to the " & TemplateName & " Installer!" & vbNewLine & vbNewLine & _
            "You need to install the newest version of the " & TemplateName & "." & vbNewLine & vbNewLine & _
            "Please click OK to begin the installation. It should only take a few seconds."
    
        If MsgBox(strWelcome, vbOKCancel, TemplateName) = vbCancel Then
            MsgBox "Please try to install the files at a later time."
            
            If Installer = True Then
                'Application.Quit (wdDoNotSaveChanges)
            End If
            
            Exit Sub
        End If
    End If
    
    ' ---------------- Close any open docs (with prompt) -----------------------------------
    Call CloseOpenDocs
        
    '----------------- download template files ------------------------------------------
    Dim logString As String
    Dim d As Long
    
    For d = LBound(strInstallMe()) To UBound(strInstallMe())
        'If False, error in download; user was notified in DownloadFromConfluence function
        If DownloadFromConfluence(FinalDir(d), strFullLogPath(d), FileName(d)) = False Then
            If Installer = True Then
                'Application.Quit (wdDoNotSaveChanges)
            Else
                Exit Sub
            End If
        End If
    Next d
    
    '------Display installation complete message and close doc (ending sub)---------------
    Dim strComplete As String
    
    strComplete = "The " & TemplateName & " has been installed on your computer." & vbNewLine & vbNewLine & _
        "When you restart Word, the template will be available."
        
    MsgBox strComplete, vbOKOnly, "Installation Successful"
    
    '------Close and restart Word for template changes to take effect---------------------
    'Would love to get restart to work...
    'Dim restartTime As Variant
    'restartTime = Now + TimeValue("00:00:01")
    'Application.OnTime When:=restartTime, Name:="Restart"
    If Installer = True Then
        'Application.Quit SaveChanges:=wdDoNotSaveChanges          'DEBUG: comment out this line
    End If
        
End Sub
Private Function DownloadFromConfluence(FinalDir As String, LogFile As String, FileName As String) As Boolean
'FinalPath is directory w/o file name

    Dim logString As String
    Dim strTmpPath As String
    Dim strBashTmp As String
    Dim strFinalPath As String
    Dim strErrMsg As String
    Dim myURL As String

    logString = ""
    strTmpPath = Environ("TEMP") & Application.PathSeparator & FileName 'Environ gives temp dir for Mac too?
    strBashTmp = Replace(strTmpPath, "\", "/")
    Debug.Print strBashTmp
    strFinalPath = FinalDir & Application.PathSeparator & FileName
    
    'this is download link, actual page housing files is https://confluence.macmillan.com/display/PBL/Test
    myURL = "https://confluence.macmillan.com/download/attachments/9044274/" & FileName
            
    #If Mac Then
        'check for network.
        If ShellAndWaitMac("ping -o google.com &> /dev/null ; echo $?") <> 0 Then   'can't connect to internet
            logString = Now & " -- Tried update; unable to connect to network."
            LogInformation LogFile, logString
            strErrMsg = "There was an error trying to download the Macmillan template." & vbNewLine & vbNewLine & _
                        "Please check your internet connection or contact workflows@macmillan.com for help."
            MsgBox strErrMsg, vbCritical, "Error 1: Connection error (" & FileName & ")"
            DownloadFromConfluence = False
            Exit Function
        Else 'internet is working, download file
            ShellAndWaitMac ("rm -f " & strBashTmp & " ; curl -o " & strBashTmp & " " & myURL)
        End If
    #Else
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

    'If final dir = Startup, disable template
    'Debug.Print strFinalPath
    If InStr(1, LCase(strFinalPath), LCase("startup"), vbTextCompare) > 0 Then         'LCase because "startup" was staying in all caps for some reason, UCase wasn't working
        On Error Resume Next                                        'Error = add-in not available, don't need to uninstall
            AddIns(strFinalPath).Installed = False
        On Error GoTo 0
    End If
    
    'If file exists already, log it and delete it
    If IsItThere(strFinalPath) = True Then
        logString = Now & " -- Previous version file in final directory."
        LogInformation LogFile, logString
        
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
        
    Else
        logString = Now & "No previous version file in final directory."
        LogInformation LogFile, logString
    End If
        
    'If delete was successful, move downloaded file to final folder
    If IsItThere(strFinalPath) = False Then
        logString = Now & " -- Final directory clear of " & FileName & " file."
        LogInformation LogFile, logString
        Name strTmpPath As strFinalPath
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
    
    DownloadFromConfluence = True

End Function

Private Sub LogInformation(LogFile As String, LogMessage As String)

Dim FileNum As Integer
    FileNum = FreeFile ' next file number
    Open LogFile For Append As #FileNum ' creates the file if it doesn't exist
    Print #FileNum, LogMessage ' write information at the end of the text file
    Close #FileNum ' close the file
End Sub

Private Function CheckLog(StyleDir As String, LogDir As String, LogPath As String) As Boolean
    
    Dim logString As String
    
    '------------------ Check log file --------------------------------------------
    'Check if logfile/directory exists
    If IsItThere(LogPath) = False Then
        CheckLog = False
        logString = Now & "----------------------------" & vbNewLine & Now & " -- Creating logfile."
        If IsItThere(LogDir) = False Then
            If IsItThere(StyleDir) = False Then
                MkDir (StyleDir)
                MkDir (LogDir)
                logString = "----------------------------" & vbNewLine & Now & " -- Creating MacmillanStyleTemplate directory."
            Else
                MkDir (LogDir)
                logString = "----------------------------" & vbNewLine & Now & " -- Creating log directory."
            End If
        End If
    Else    'logfile exists, so check last modified date
        Dim lastModDate As Date
        lastModDate = FileDateTime(LogPath)
        If DateDiff("d", lastModDate, Date) < 1 Then       'i.e. 1 day
            CheckLog = True
            logString = "----------------------------" & vbNewLine & Now & " -- Already checked less than 1 day ago."
        Else
            CheckLog = False
            logString = "----------------------------" & vbNewLine & Now & " -- >= 1 day since last update check."
        End If
    End If
    
    'Log that info!
    LogInformation LogPath, logString
    
End Function

Private Function IsTemplateThere(Directory As String, FileName As String, Log As String)
    
    'Create full path to template
    Dim strFullPath As String
    Dim logString As String
    strFullPath = Directory & Application.PathSeparator & FileName
    
    '------------------------- Check if template exists ----------------------------------
    ' Create directory if it doesn't exist
    If IsItThere(Directory) = False Then
        MkDir (Directory)
        IsTemplateThere = False
        logString = Now & " -- Creating template directory."
    Else
        ' Check if template file exists
        If IsItThere(strFullPath) = False Then
            IsTemplateThere = False
            logString = Now & " -- " & FileName & " doesn't exist in " & Directory
        Else
            IsTemplateThere = True
            logString = Now & " -- " & FileName & " already exists."
        End If
    End If

    LogInformation Log, logString
End Function

Private Function NeedUpdate(Directory As String, FileName As String, Log As String) As Boolean
'Directory argument should be the final directory the template should go in.
'File should be the template file name
'Log argument should be full path to log file

    '------------------------- Get installed version number -----------------------------------
    Dim logString As String
    Dim strFullTemplatePath As String
    
    strFullTemplatePath = Directory & Application.PathSeparator & FileName
    'Debug.Print "NeedUpdate Path: " & strFullTemplatePath
    
    'Get version number of installed template
    Dim strInstalledVersion As String
    If IsItThere(strFullTemplatePath) = True Then
        Documents.Open FileName:=strFullTemplatePath, ReadOnly:=True, Visible:=False
        strInstalledVersion = Documents(strFullTemplatePath).CustomDocumentProperties("Version")
        Documents(strFullTemplatePath).Close
        logString = Now & " -- installed version is " & strInstalledVersion
    Else
        strInstalledVersion = 0     ' Template is not installed
        logString = Now & " -- No template installed, version number is 0."
    End If
    
    LogInformation Log, logString
    
    '------------------------- Try to get current version's number from Confluence ------------
    Dim strVersion As String
    Dim strFullVersionPath As String
    
    'Debug.Print FileName
    'Debug.Print InStrRev(FileName, ".do")
    strVersion = Left(FileName, InStrRev(FileName, ".do") - 1)
    strVersion = strVersion & ".txt"
    strFullVersionPath = Directory & Application.PathSeparator & strVersion
    'Debug.Print strVersion
    
    'If False, error in download; user was notified in DownloadFromConfluence function
    If DownloadFromConfluence(Directory, Log, strVersion) = False Then
        NeedUpdate = False
    End If
        
    '-------------------- Get version number of current template ---------------------
    If IsItThere(strFullVersionPath) = True Then
        NeedUpdate = True
        Dim strCurrentVersion As String
        strCurrentVersion = ImportVariable(strFullTemplatePath)
        logString = Now & " -- Current version is " & strCurrentVersion
    Else
        NeedUpdate = False
        logString = Now & " -- Download of version file for " & FileName & " failed."
    End If
        
    LogInformation Log, logString
    
    If NeedUpdate = False Then
        Exit Function
    End If
    
    '--------------------- Compare version numbers -----------------------------------
    
    If strInstalledVersion >= strCurrentVersion Then
        NeedUpdate = False
        logString = Now & " -- Current version matches installed version."
    Else
        NeedUpdate = True
        logString = Now & " -- Current version greater than installed version."
    End If
    
    LogInformation Log, logString
    
End Function

Private Function IsItThere(Path)
' Check if file or directory exists
    
    Debug.Print Path
    
    'Remove trailing path separator from dir if it's there
    If Right(Path, 1) = Application.PathSeparator Then
        Path = Left(Path, Len(Path) - 1)
    End If
    
    Dim CheckDir As String
    Dim lngAttributes As Long
    On Error GoTo ErrHandler            ' Because Dir(Path) throws an error on Mac if not existant
    
    'Includes checks for read-only, hidden, or system files, or for directories
    'lngAttributes = (vbReadOnly Or vbHidden Or vbSystem Or vbDirectory)
    
    CheckDir = Dir(Path, vbDirectory)
    
    If CheckDir = vbNullString Then
        IsItThere = False
    Else
        IsItThere = True
    End If
    
Exit Function

ErrHandler:
    If Err.Number = 68 Then     ' "Device unavailable"
        IsItThere = False
    Else
        Debug.Print Err.Number & ": " & Err.Description
    End If
End Function

Private Sub CloseOpenDocs()

    '-------------Check for/close open documents---------------------------------------------
    Dim strInstallerName As String
    Dim strSaveWarning As String
    Dim objDocument As Document
    Dim b As Long
    
    strInstallerName = ThisDocument.Name
        'Debug.Print "Installer Name: " & strInstallerName
        'Debug.Print "Open docs: " & Documents.Count
        
    If Documents.Count > 1 Then
        strSaveWarning = "All other Word documents must be closed to run the installer." & vbNewLine & vbNewLine & _
            "Click OK and I will save and close your documents." & vbNewLine & _
            "Click Cancel to exit without installing and close the documents yourself."
        If MsgBox(strSaveWarning, vbOKCancel, "Close documents?") = vbCancel Then
            ActiveDocument.Close
            Exit Sub
        Else
            For b = 1 To Documents.Count
                'Debug.Print "Current doc " & b & ": " & Documents(b).Name
                On Error Resume Next        'To skip error if user is prompted to save new doc and clicks Cancel
                    If Documents(b).Name <> strInstallerName Then       'But don't close THIS document
                        Documents(b).Save   'separate step to trigger Save As prompt for previously unsaved docs
                        Documents(b).Close
                    End If
                On Error GoTo 0
            Next b
        End If
    End If
    
End Sub

Private Function ImportVariable(strFile As String) As String
 
    Open strFile For Input As #1
    Line Input #1, ImportVariable
    Close #1
 
End Function

Private Function IsArrayEmpty(Arr As Variant) As Boolean
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

Private Function ShellAndWaitMac(cmd As String) As String

    Dim result As String
    Dim scriptCmd As String ' Macscript command
    
    scriptCmd = "do shell script """ & cmd & """"
    result = MacScript(scriptCmd) ' result contains stdout, should you care
    ShellAndWaitMac = result

End Function
