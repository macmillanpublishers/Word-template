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
'If this is an installer file, The Part 1 code needs to reside in the ThisDocument module of a .docm file so that
'it will launch when users open the file.

'"Installer" argument = True if this is for a standalone installtion file.
'"Installer" argument = False is this is part of a daily check of the current file and only update if out of date.
    
    '' --------------- Check that variables were passed correctly -------------------------
    'Dim x As Long
    'For x = LBound(FileName()) To UBound(FileName())
    '    Debug.Print & " " & FileName(x) & " " & FinalDir(x) & vbNewLine
    'Next x
    
    '' Create logfile names and check if need update (for updater)
    Dim a As Long
    Dim strLogFile() As String
    ReDim strLogFile(LBound(FileName()) To UBound(FileName()))
    Dim installCheck() As Boolean
    ReDim installCheck(LBound(FileName()) To UBound(FileName()))
    Dim strFullLogPath() As String
    ReDim strFullLogPath(LBound(FileName()) To UBound(FileName()))

    For a = LBound(FileName()) To UBound(FileName())
        'Create logfile name
        strLogFile(a) = Left(FileName(a), InStrRev(FileName(a), ".do") - 1)
        #If Mac Then
            Dim strUser As String
            strUser = MacScript("tell application " & Chr(34) & "System Events" & Chr(34) & Chr(13) & _
                "return (name of current user)" & Chr(13) & "end tell")
            strFullLogPath(a) = "Macintosh HD:Users:" & strUser & ":Documents:MacmillanStyleTemplate:log:" & strLogFile(a) & "_updates.log"
        #Else
            strFullLogPath(a) = Environ("ProgramData") & "\MacmillanStyleTemplate\log\" & strLogFile(a) & "_updates.log"
        #End If
        'Debug.Print strFullLogPath(a)
        
        'If this is an updated (not an installer) check if has been checked today, if no check version number
        If Installer = False Then
            installCheck(a) = NeedUpdate(Directory:=FinalDir(a), File:=FileName(a), Log:=strLogFile(a))
        End If
    Next a
    
    'If all installCheck are false, Exit Sub
    'Else, create new array for file names and directories (or combine?) for just installCheck=true
    
    
    
    
    
    
    
    ' ---------------- Alert user that installation is happening --------------------------
    Dim strWelcome As String

    strWelcome = "Welcome to the " & TemplateName & " Installer!" & vbNewLine & vbNewLine & _
        "You need to install the newest version of the " & TemplateName & "." & vbNewLine & vbNewLine & _
        "Please click OK to begin the installation. It should only take a few seconds."

    If MsgBox(strWelcome, vbOKCancel, TemplateName) = vbCancel Then
        ActiveDocument.Close (wdDoNotSaveChanges)
        Exit Sub
    End If
    
    Call CloseOpenDocs
        
    '-----------------Define variables--------------------------------------------------
    'For template file
    Dim strStartupDir As String
    Dim strGtFinalPath As String
    
    'For style templates & log file
    Dim strStyleDir As String
    
    'For log files
    Dim strLogDir As String
    'Dim strLogFile As String
    Dim strLogPath As String
    Dim logString As String

    strStartupDir = Application.StartupPath
    strGtFinalPath = strStartupDir & "\" & FileName(1)
    
    strStyleDir = Environ("PROGRAMDATA") & "\MacmillanStyleTemplate"
    
    strLogDir = strStyleDir & "\log"
    strLogFile = "mac_templates_" & Format(Date, "yyyy-mm-dd") & "_" & Format(Time, "hh-mm-ss") & ".log"
    strLogPath = strLogDir & "\" & strLogFile(1)
    logString = ""
        
    '----------------Check for and create log file----------------------------------------
    If Dir(strLogDir, vbDirectory) <> vbNullString Then                 'If log dir already exists
        logString = "-- log directory already exists."
    Else
        If Dir(strStyleDir, vbDirectory) = vbNullString Then            'If MacmillanStyleTemplate dir doesn't exist, then create
            MkDir (strStyleDir)
            MkDir (strLogDir)
            logString = "-- created MacmillanStyleTemplate directory and log file."
        Else                                                            'MacStyleTemplate exists but log dir doesn't
            MkDir (strLogDir)
            logString = "-- created log directory and log file."
        End If
    End If
    
    'write logString to log file
    LogInformation strLogPath, logString
        
    '-------------download template file------------------------------------------
    'Log attempt to download
    logString = "------------------------------------------" & vbNewLine & _
                "DOWNLOAD " & FileName(1) & vbNewLine & _
                "------------------------------------------"
    LogInformation strLogPath, logString
    
    'If False, error in download; user was notified in DownloadFromConfluence function
    If DownloadFromConfluence(strGtFinalPath, strLogPath, FileName(1)) = False Then
        ActiveDocument.Quit (wdDoNotSaveChanges)
        Exit Sub
    End If
    
    '------Display installation complete message and close doc (ending sub)---------------
    Dim strComplete As String
    
    strComplete = "The " & TemplateName & " has been installed on your computer." & vbNewLine & vbNewLine & _
        "Click OK to close Word. When you restart Word, the template will be available."
        
    MsgBox strComplete, vbOKOnly, "Installation Successful"
    
    '------Close and restart Word for template changes to take effect---------------------
    'Would love to get restart to work...
    'Dim restartTime As Variant
    'restartTime = Now + TimeValue("00:00:01")
    'Application.OnTime When:=restartTime, Name:="Restart"
    Application.Quit SaveChanges:=wdDoNotSaveChanges          'DEBUG: comment out this line
    
End Sub
Private Function DownloadFromConfluence(FinalPath As String, LogFile As String, FileName As String) As Boolean
        
    Dim logString As String
    Dim strTmpPath As String
    Dim strErrMsg As String
        
    logString = ""
    strTmpPath = Environ("TEMP") & "\" & FileName
        
    'try to download the file from Public Confluence page
    Dim myURL As String
    Dim WinHttpReq As Object
    Dim oStream As Object
    
    'this is download link, actual page housing files is https://confluence.macmillan.com/display/PBL/Test
    myURL = "https://confluence.macmillan.com/download/attachments/9044274/" & FileName
        
    'Attempt to download file
    On Error Resume Next
        Set WinHttpReq = CreateObject("MSXML2.XMLHTTP.3.0")
        WinHttpReq.Open "GET", myURL, False
        WinHttpReq.Send

            ' Exit sub if error in connecting to website
            If Err.Number <> 0 Then 'HTTP request is not OK
                'Debug.Print WinHttpReq.Status
                logString = "-- could not connect to Confluence site: Error " & Err.Number & ". Exiting installation."
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
        logString = "-- Http status is " & WinHttpReq.Status & ". Cannot download file, exiting installer."
        LogInformation LogFile, logString
        strErrMsg = "There was an error trying to download the Macmillan templates." & vbNewLine & vbNewLine & _
            "Please check your internet connection or contact workflows@macmillan.com for help."
        MsgBox strErrMsg, vbCritical, "Error 2: Http status " & WinHttpReq.Status & " (" & FileName & ")"
        DownloadFromConfluence = False
        Exit Function
    End If
        
    'Error if download was not successful
    If Dir(strTmpPath) = vbNullString Then
        logString = "-- " & FileName & " file download to Temp was not successful. Exiting installer."
        LogInformation LogFile, logString
        strErrMsg = "There was an error downloading the Macmillan template." & vbNewLine & _
            "Please contact workflows@macmillan.com for assitance."
        MsgBox strErrMsg, vbCritical, "Error 3: Download failed (" & FileName & ")"
        DownloadFromConfluence = False
        Exit Function
    Else
        logString = "-- " & FileName & " file download to Temp was successful."
        LogInformation LogFile, logString
    End If

    'If final dir = Startup, disable template
    If InStr(LCase(FinalPath), LCase("startup")) > 0 Then           'LCase cause startup was staying in all caps for some reason
        On Error Resume Next                                        'Error = add-in not available, don't need to uninstall
            AddIns(FinalPath).Installed = False
        On Error GoTo 0
    End If
    
    'If file exists already, log it and delete it
    If Dir(FinalPath) <> vbNullString Then
        logString = "-- Previous version file in final directory."
        LogInformation LogFile, logString
        
        On Error Resume Next
            Kill FinalPath
            
            If Err.Number = 70 Then         'File is open and can't be replaced
                logString = "-- old " & FileName & " file is open, can't delete/replace. Alerting user, exiting sub."
                LogInformation LogFile, logString
                strErrMsg = "Please close all other Word documents and try again."
                MsgBox strErrMsg, vbCritical, "Error 4: Previous version removal failed (" & FileName & ")"
                DownloadFromConfluence = False
                On Error GoTo 0
                Exit Function
            End If
        On Error GoTo 0
        
    Else
        logString = "No previous version file in final directory."
        LogInformation LogFile, logString
    End If
        
    'If delete was successful, move downloaded file to Startup folder
    If Dir(FinalPath) = vbNullString Then
        logString = "-- Final directory clear of " & FileName & " file."
        LogInformation LogFile, logString
        Name strTmpPath As FinalPath
    Else
        logString = "-- old " & FileName & " file not cleared from Final directory. Exiting installer."
        LogInformation LogFile, logString
        strErrMsg = "There was an error installing the Macmillan template." & vbNewLine & _
            "Please close all other Word documents and try again, or contact workflows@macmillan.com."
        MsgBox strErrMsg, vbCritical, "Error 5: Previous version uninstall failed (" & FileName & ")"
        DownloadFromConfluence = False
        Exit Function
    End If
    
    'If move was successful, yay! Else, :(
    If Dir(FinalPath) <> vbNullString Then
        logString = "-- " & FileName & " file successfully saved to final directory."
        LogInformation LogFile, logString
    Else
        logString = "-- " & FileName & " file not saved to final directory."
        LogInformation LogFile, logString
        strErrMsg = "There was an error installing the Macmillan template." & vbNewLine & vbNewLine & _
            "Please cotact workflows@macmillan.com for assistance."
        MsgBox strErrMsg, vbCritical, "Error 6: Installation failed (" & FileName & ")"
        DownloadFromConfluence = False
        Exit Function
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

Private Function NeedUpdate(Directory As String, File As String, Log As String) As Boolean
    
    '------------------- Create full path to template file ----------------------------------
    'Remove trailing path separator from dir if it's there, so we know we won't duplicate below
    If Right(Directory, 1) = Application.PathSeparator Then
        Directory = Left(Directory, Len(Directory) - 1)
    End If
    
    Dim strTemplatePath As String
    strTemplatePath = Directory & Application.PathSeparator & File
    
    '------------------- Create path to log file --------------------------------------------
    Dim strStyleDir As String
    Dim strLogDir As String
    Dim strLogPath As String
    
    #If Mac Then
        Dim strUser As String
        strUser = MacScript("tell application " & Chr(34) & "System Events" & Chr(34) & Chr(13) & _
                "return (name of current user)" & Chr(13) & "end tell")
        strStyleDir = "Macintosh HD:Users:" & strUser & ":Documents:MacmillanStyleTemplate"
    #Else
        strStyleDir = Environ("ProgramData") & "\MacmillanStyleTemplate"
    #End If
    
    strLogDir = strStyleDir & Application.PathSeparator & "log"
    strLogPath = strLogDir & Application.PathSeparator & Log
    
    Dim logString As String
    
    '------------------ Check log file --------------------------------------------
    'Check if logfile/directory exists
    If IsItThere(strLogPath) = False Then
        NeedUpdate = True
        logString = Now & "-- Creating logfile."
        If IsItThere(strLogDir) = False Then
            If IsItThere(strStyleDir) = False Then
                MkDir (strStyleDir)
                MkDir (strLogDir)
                logString = Now & "-- Creating MacmillanStyleTemplate directory."
            Else
                MkDir (strLogDir)
                logString = Now & "-- Creating log directory."
            End If
        End If
    Else    'logfile exists, so check last modified date
        Dim lastModDate As Date
        lastModDate = FileDateTime(strLogPath)
        If DateDiff("d", lastModDate, Date) < 1 Then       'i.e. 1 day
            NeedUpdate = False
            logString = Now & "-- Already checked less than 1 day ago. Exiting updater."
        Else
            NeedUpdate = True
            logString = Now & "-- >= 1 day since last update check."
        End If
    End If
    
    'Log that info!
    LogInformation strLogPath, logString
    
    '===========================
    ' Debugging: Set to True
    NeedUpdate = True
    '===========================
    
    '------------------------- Check if template exists ----------------------------------
    'If we already checked today, don't need to update.
    If NeedUpdate = False Then
        Exit Function
    Else ' Let's check the version number
        ' Create directory if it doesn't exist
        If IsItThere(Directory) = False Then
            MkDir (Directory)
            NeedUpdate = True
            logString = Now & "-- Creating template directory."
        Else
            ' Check if template file exists
            If IsItThere(strTemplatePath) = False Then
                NeedUpdate = True
                logString = Now & "-- " & File & " doesn't exist in " & Directory
            Else
                NeedUpdate = False
                logString = Now & "-- " & File & " already exists. Checking version number."
            End If
    End If

    LogInformation strLogPath, logString
    
    '------------------------- Get installed version number -----------------------------------
    If NeedUpdate = True Then
        Exit Function
    Else
        'Get version number of installed template
        Dim strInstalledVersion As String
        Documents.Open FileName:=strTemplatePath, ReadOnly:=True, Visible:=False
        strInstalledVersion = Documents(strTemplatePath).CustomDocumentProperties("version")
        Documents(strTemplatePath).Close
        logString = Now & "-- installed version is " & strInstalledVersion
    End If
    
    LogInformation strLogPath, logString
    
    '------------------------- Try to get current version's number from Confluence ------------
    Dim strVersion As String
    Dim strTempPath As String
        
    strVersion = Left(File(, InStrRev(FileName), ".do") - 1)
    strVersion = strVersion & ".txt"
    strTempPath = Environ("TEMP") & Application.PathSeparator & strVersion
    
    #If Mac Then
        'Mac way of getting
    #Else
        Dim myURL As String
        Dim WinHttpReq As Object
        Dim oStream As Object
        Dim templateURL As String
        
        'this is download link, actual page housing template is http://confluence.macmillan.com/display/PBL/Test
        templateURL = "https://confluence.macmillan.com/download/attachments/9044274/"
        myURL = templateURL & strVersion
    
        Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
        WinHttpReq.Open "GET", myURL, False
        
        On Error Resume Next
        WinHttpReq.Send
    
            ' Exit function if error in connecting to website
            If Err.Number <> 0 Then 'HTTP request is not OK
                logString = Now & " -- tried to update " & templateFile & "; unable to connect to Confluence website."
                LogInformation strLogPath, logString
                NeedUpdate = False 'can't update template :(
                Exit Function
            End If
        
        On Error GoTo 0
    
        Debug.Print WinHttpReq.Status
    
        If WinHttpReq.Status = 200 Then  ' 200 = HTTP request is OK
        
            'if connection OK, download file and save to temp directory
            myURL = WinHttpReq.responseBody
            Set oStream = CreateObject("ADODB.Stream")
            oStream.Open
            oStream.Type = 1
            oStream.Write WinHttpReq.responseBody
            oStream.SaveToFile strTempPath, 2 ' 1 = no overwrite, 2 = overwrite
            oStream.Close
            Set oStream = Nothing
            Set WinHttpReq = Nothing
        End If
    #End If
        
    '-------------------- Get version number of current template ---------------------
    If IsItThere(strTempPath) = True Then
        Dim strCurrentVersion As String
        strCurrentVersion = ImportVariable(strTempPath)
        logString = Now & "current version is " & strCurrentVersion
    Else
        NeedUpdate = False
        logString = Now & "-- download of version file for " & File & " failed."
    End If
        
    LogInformation strLogPath, logString
    
    '--------------------- Compare version numbers -----------------------------------
    
    If strInstalledVersion >= strCurrentVersion Then
        NeedUpdate = False
        logString = Now & "-- Current version matches installed version."
    Else
        NeedUpdate = True
        logString = Now & "-- Current version greater than installed version."
    End If
    
End Function

Private Function IsItThere(Path)
' Check if file or directory exists
' Need error handler because on Mac 2011 if file/dir doesn't exist, it throws an error
' If checking for directory, path must include trailing separator
    
    'Remove trailing path separator from dir if it's there, so we know we won't duplicate below
    If Right(Path, 1) = Application.PathSeparator Then
        Path = Left(Path, Len(Path) - 1)
    End If

    'Now that we know there ISN'T a trailing separator, we'll add it so we know there is one there
    Path = Path & Application.PathSeparator
    
    Dim CheckDir As String
    On Error GoTo ErrHandler
    
    CheckDir = Dir(Path)
    
    If CheckDir = vbNullString Then
        IsItThere = False
    Else
        IsItThere = True
    End If
    
Exit Function

ErrHandler:
    If Err.Number = 68 Then     ' "Device unavailable"
        IsItThere = False
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
