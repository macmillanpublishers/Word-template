Attribute VB_Name = "SharedFileInstaller"
' created by Erica Warren - erica.warren@macmillan.com

' ======== PURPOSE ===================================
' Downloads and installs an array of template files & logs the downloads

' ======== DEPENDENCIES =======================================
' This is Part 2 of 2. Must be called from a sub in another module that declares file names and locations.
' The template file needs to be uploaded as an attachment to https://confluence.macmillan.com/display/PBL/Test
' If this is an installer file, The Part 1 code needs to reside in the ThisDocument module as a sub called
' Documents_Open in a .docm file so that it will launch when users open the file.
' Requires the SharedMacros module be installed in the same template.

Option Explicit
Option Base 1

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

Public Sub Installer(DownloadFrom As GitBranch, Installer As Boolean, TemplateName As String, ByRef TemplatesToInstall() As String)

'"Installer" argument = True if this is for a standalone installtion file.
'"Installer" argument = False is this is part of a daily check of the current file and only updates if out of date.
    
    ' Separate file name from directory path
    Dim lngBreak As Long
    Dim FileName() As String
    Dim FinalDir() As String
    Dim Z As Long
    
    For Z = LBound(TemplatesToInstall()) To UBound(TemplatesToInstall())
'        Debug.Print "Path: " & TemplatesToInstall(z)
        
        lngBreak = InStrRev(TemplatesToInstall(Z), Application.PathSeparator)
'        Debug.Print "Final sep at: " & lngBreak
        
        If lngBreak >= 1 Then
            ReDim Preserve FileName(1 To Z)
            FileName(Z) = Right(TemplatesToInstall(Z), (Len(TemplatesToInstall(Z)) - lngBreak))
            
            ReDim Preserve FinalDir(1 To Z)
            FinalDir(Z) = Left(TemplatesToInstall(Z), lngBreak - 1)
            
'            Debug.Print "File Name #" & z & ": " & FileName(z)
'            Debug.Print "Directory #" & z & ": " & FinalDir(z)
        Else
            ' No path separator in full path specification
            MsgBox "You need to specify the full path to your templates!"
            Exit Sub
        End If
        
    Next Z
    
    
    '' --------------- Set up variable names ----------------------------------------------
    '' Create style directory and logfile names
    Dim A As Long
    Dim arrLogInfo() As Variant
    ReDim arrLogInfo(1 To 3)
    Dim strStyleDir() As String
    ReDim strStyleDir(LBound(FileName()) To UBound(FileName()))
    Dim strLogDir() As String
    ReDim strLogDir(LBound(FileName()) To UBound(FileName()))
    Dim strFullLogPath() As String
    ReDim strFullLogPath(LBound(FileName()) To UBound(FileName()))
    
    ' ------------ Define Log Dirs and such -----------------------------------------
    For A = LBound(FileName()) To UBound(FileName())
        arrLogInfo() = CreateLogFileInfo(FileName(A))
        strStyleDir(A) = arrLogInfo(1)
        strLogDir(A) = arrLogInfo(2)
        strFullLogPath(A) = arrLogInfo(3)
    Next A
    
    'Debug.Print "Style Dir is: " & strStyleDir(1) & vbNewLine & _
                "Log dir is: " & strLogDir(1) & vbNewLine & _
                "Full path to log is: " & strFullLogPath(1)
                
    ' ------------- Check if we need to do an installation ---------------------------
    ' Check if template exists
    Dim installCheck() As Boolean
    ReDim installCheck(LBound(FileName()) To UBound(FileName()))
    Dim blnTemplateExists() As Boolean
    ReDim blnTemplateExists(LBound(FileName()) To UBound(FileName()))
    Dim blnLogUpToDate() As Boolean
    ReDim blnLogUpToDate(LBound(FileName()) To UBound(FileName()))
    Dim logString As String
    Dim strTypeOfInstall As String

    Dim B As Long
    
    For B = LBound(FileName()) To UBound(FileName())
        
        ' Check if log dir/file exists, create if it doesn't, check last mod date if it does
        ' We don't need the true/false info for Installer, but we DO need to run these two
        ' functions to create directories if they don't exist yet
        
        ' If last mod date less than 1 day ago, CheckLog = True
        blnLogUpToDate(B) = CheckLog(strStyleDir(B), strLogDir(B), strFullLogPath(B))
        'Debug.Print FileName(b) & " log exists and was checked today: " & blnLogUpToDate(b)
        
        ' Check if template exists, if not create any missing directories
        blnTemplateExists(B) = IsTemplateThere(FinalDir(B), FileName(B), strFullLogPath(B))
        ' Debug.Print FileName(b) & " exists: " & blnTemplateExists(b)
        
        ' ===============================
        ' FOR DEBUGGING: SET TO TRUE,    |
        ' SO ALWAYS DOWNLOADS FILES      |
        ' Installer = True              '|
        ' ===============================
        
        If Installer = False Then 'Because if it's an installer, we just want to install the file

                
            ' ==========================================
            ' FOR DEBUGGING: SET TO FALSE AND THEN TRUE |
            ' TO TEST NEEDUPDATE FUNCTION               |
            ' blnLogUpToDate(b) = False                '|
            ' blnTemplateExists(b) = True              '|
            ' ==========================================
                
            If blnLogUpToDate(B) = True And blnTemplateExists(B) = True Then ' already checked today, already exists
                installCheck(B) = False
            ElseIf blnLogUpToDate(B) = False And blnTemplateExists(B) = True Then 'Log is new or not checked today, already exists
                'check version number
                installCheck(B) = NeedUpdate(DownloadFrom, FinalDir(B), FileName(B), strFullLogPath(B))
            Else ' blnTemplateExists = False, just download new template
                 installCheck(B) = True
            End If
        Else
            installCheck(B) = True
        End If
        
    Next B

    ' ---------------- Create new array of template files we need to install -----------------
    Dim strInstallFile() As String
    Dim strInstallDir() As String
    Dim C As Long
    Dim X As Long
    
    X = 0
    
    For C = LBound(FileName()) To UBound(FileName())
        If installCheck(C) = True Then
            X = X + 1
            ReDim Preserve strInstallFile(1 To X)
                strInstallFile(X) = FileName(C)
            ReDim Preserve strInstallDir(1 To X)
                strInstallDir(X) = FinalDir(C)
        End If
    Next C
    
    'Debug.Print strInstallFile(1) & vbNewLine & strInstallDir(1)
    
    ' ---------------- Check if new array is allocated -----------------------------------
    If IsArrayEmpty(strInstallFile()) = True Then       ' No files need to be installed
        If Installer = True Then  ' Though this option (no files to install on installer) shouldn't actually occur
            #If Mac Then    ' because application.quit generates error on Mac
                activeDoc.Close (wdDoNotSaveChanges)
            #Else
                Application.Quit (wdDoNotSaveChanges)
            #End If
        Else
            Exit Sub
        End If
    Else ' There are values in the array and we need to install them
    
        ' Alert user that installation is happening
        Dim strWelcome As String
        If Installer = True Then
            strWelcome = "Welcome to the " & TemplateName & " Installer!" & vbNewLine & vbNewLine & _
                "Please click OK to begin the installation. It should only take a few seconds."
        Else
            strWelcome = "Your " & TemplateName & " is out of date. Click OK to update automatically."
        End If
    
        If MsgBox(strWelcome, vbOKCancel, TemplateName) = vbCancel Then
            MsgBox "Please try to install the files at a later time."
            
            If Installer = True Then
                activeDoc.Close (wdDoNotSaveChanges)
            End If
            
            Exit Sub
        End If
    End If
    
    ' ---------------- Close any open docs (with prompt) -----------------------------------
    Call CloseOpenDocs
        
    '----------------- download template files ------------------------------------------
    Dim D As Long
    
    For D = LBound(strInstallFile()) To UBound(strInstallFile())
    
        If IsReadOnly(strInstallDir(D)) = True Then
            ' Can't replace with new file if destination is read-only; Startup on Mac w/o admin is read-only
            Dim strReadOnlyError As String
            
            strReadOnlyError = "Sorry, you don't have permission to install the file " & strInstallFile(D) & vbNewLine & vbNewLine & _
                "If you are in-house at Macmillan on a Mac, try re-installing the Macmillan Style Template & Macros from the Digital Workflow category in Self Service."
                
                MsgBox strReadOnlyError, vbOKOnly, "Update Failed"
                Exit Sub
        Else
            'If False, error in download; user was notified in DownloadFromConfluence function
            If DownloadFromConfluence(DownloadSource:=DownloadFrom, FinalDir:=strInstallDir(D), _
                LogFile:=strFullLogPath(D), FileName:=strInstallFile(D)) = False Then
                If Installer = True Then
                    #If Mac Then    ' because application.quit generates error on Mac
                        activeDoc.Close (wdDoNotSaveChanges)
                    #Else
                        Application.Quit (wdDoNotSaveChanges)
                    #End If
                Else
                    Exit Sub
                End If
            End If
        End If
        
        ' If we just updated the main template, delete the old toolbar
        ' Will be added again by Word-template AutoExec when it's launched, to capture updates
        #If Mac Then
            Dim Bar As CommandBar
            If strInstallFile(D) = "Word-template.dotm" Then
                For Each Bar In CommandBars
                    If Bar.Name = "Macmillan Tools" Then
                        Bar.Delete
                        'Exit For  ' Actually don't exit, in case there are multiple toolbars
                    End If
                    Next
            End If
        #End If
    Next D
    
    '------Display installation complete message   ---------------------------
    Dim strComplete As String
    Dim strInstallType As String
    
    ' Quit if it's an installer, but not if it's an updater (updater was causing conflicts between GT and GtUpdater)
    If Installer = True Then
        strInstallType = "installed"
    Else
        strInstallType = "updated"
    End If
    
    strComplete = "The " & TemplateName & " has been " & strInstallType & " on your computer."
    MsgBox strComplete, vbOKOnly, "Installation Successful"
    
End Sub

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


Public Function DownloadCSV(FileName As String, Optional DownloadFrom As GitBranch = master) As Variant
    '---------Download CSV with design specs from Confluence site-------

    'Create log file name
    Dim arrLogInfo() As Variant
    ReDim arrLogInfo(1 To 3)
    
    arrLogInfo() = CreateLogFileInfo(FileName)
      
    'Create final path for downloaded CSV file (in log directory)
    'not in temp dir because that is where DownloadFromConfluence downloads it to, and it cleans that file up when done
    Dim strStyleDir As String
    Dim strPath As String
    Dim strLogFile As String
    Dim strMessage As String
    Dim strDir As String
    
    strStyleDir = arrLogInfo(1)
    strDir = arrLogInfo(2)
    strLogFile = arrLogInfo(3)
    strPath = strDir & Application.PathSeparator & FileName
        
    'Check if log file already exists; if not, create it
    CheckLog strStyleDir, strDir, strLogFile
    
    'Download CSV file from Confluence
    If DownloadFromConfluence(FinalDir:=strDir, LogFile:=strLogFile, FileName:=FileName, DownloadSource:=DownloadFrom) = False Then
        ' If download fails, check if we have an older version of the CSV to work with
        If IsItThere(strPath) = False Then
            strMessage = "Looks like we can't download the design info from the internet right now. " & _
                "Please check your internet connection, or contact workflows@macmillan.com."
            MsgBox strMessage, vbCritical, "Error 5: Download failed, no previous design file available"
            Exit Function
        Else
            strMessage = "Looks like we can't download the most up-to-date design info from the internet right now, " & _
                "so we'll just use the info we have on file for your castoff."
            MsgBox strMessage, vbInformation, "Let's do this thing!"
        End If
    End If
    
    ' Heading row/col different based on different InfoTypes
    Dim blnRemoveHeaderRow As Boolean
    Dim blnRemoveHeaderCol As Boolean
    
    ' Because the castoff CSV has header row and col, but Spine CSV only has a header row
    If InStr(1, FileName, "Castoff") <> 0 Then
        blnRemoveHeaderRow = True
        blnRemoveHeaderCol = True
    ElseIf InStr(1, FileName, "Spine") <> 0 Then
        blnRemoveHeaderRow = True
        blnRemoveHeaderCol = False
    ElseIf InStr(1, FileName, "Styles_Bookmaker") <> 0 Then
        blnRemoveHeaderRow = True
        blnRemoveHeaderCol = False
    End If
    
    'Double check that CSV is there
    Dim arrFinal() As Variant
    If IsItThere(strPath) = False Then
        strMessage = "The macro is unable to access the data file right now. Please check your internet " & _
                    "connection and try again, or contact workflows@macmillan.com."
        MsgBox strMessage, vbCritical, "Error 3: CSV doesn't exist"
        Exit Function
    Else
        ' Load CSV into an array
        arrFinal = LoadCSVtoArray(Path:=strPath, RemoveHeaderRow:=blnRemoveHeaderRow, RemoveHeaderCol:=blnRemoveHeaderCol)
    End If
    
    DownloadCSV = arrFinal
    
End Function


' ===== DownloadFromConfluence ================================================
' Actually now it downloads from Github but don't want to mess with things, we're
' going to be totally refacroting soon.
Private Function DownloadFromConfluence(FinalDir As String, LogFile As String, FileName As String, _
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
    ElseIf InStr(1, FileName, "macmillan", vbTextCompare) Then
      strDownloadRepo = "Word-template_assets/"
      strSubfolder = "StyleTemplate_auto-generate/"
    Else
      strDownloadRepo = "bookmaker_validator/"
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


Private Sub LogInformation(LogFile As String, LogMessage As String)

Dim FileNum As Integer
    FileNum = FreeFile ' next file number
    Open LogFile For Append As #FileNum ' creates the file if it doesn't exist
    Print #FileNum, LogMessage ' write information at the end of the text file
    Close #FileNum ' close the file
End Sub



Private Function LoadCSVtoArray(Path As String, RemoveHeaderRow As Boolean, RemoveHeaderCol As Boolean) As Variant

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



Private Function CheckLog(StylesDir As String, LogDir As String, LogPath As String) As Boolean
'LogPath is *full* path to log file, including file name. Created by CreateLogFileInfo sub, to be called before this one.

    Dim logString As String
    
    '------------------ Check log file --------------------------------------------
    'Check if logfile/directory exists
    If IsItThere(LogPath) = False Then
        CheckLog = False
        logString = Now & " -- Creating logfile."
        If IsItThere(LogDir) = False Then
            If IsItThere(StylesDir) = False Then
                MkDir (StylesDir)
                MkDir (LogDir)
                logString = Now & " -- Creating MacmillanStyleTemplate directory."
            Else
                MkDir (LogDir)
                logString = Now & " -- Creating log directory."
            End If
        End If
    Else    'logfile exists, so check last modified date
        Dim lastModDate As Date
        lastModDate = FileDateTime(LogPath)
        If DateDiff("d", lastModDate, Date) < 1 Then       'i.e. 1 day
            CheckLog = True
            logString = Now & " -- Already checked less than 1 day ago."
        Else
            CheckLog = False
            logString = Now & " -- >= 1 day since last update check."
        End If
    End If
    
    'Log that info!
    LogInformation LogPath, logString
    
End Function

Private Function CreateLogFileInfo(ByRef FileName As String) As Variant
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

Private Function NeedUpdate(DownloadURL As GitBranch, Directory As String, FileName As String, Log As String) As Boolean
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
        
        #If Mac Then
            Call OpenDocMac(strFullTemplatePath)
        #Else
            Call OpenDocPC(strFullTemplatePath)
        #End If
        
        strInstalledVersion = Documents(strFullTemplatePath).CustomDocumentProperties("Version")
        Documents(strFullTemplatePath).Close SaveChanges:=wdDoNotSaveChanges
        logString = Now & " -- installed version is " & strInstalledVersion
'        Debug.Print "InstalledVersion : |" & strInstalledVersion; "|"
    Else
        strInstalledVersion = "0"     ' Template is not installed
        logString = Now & " -- No template installed, version number is 0."
    End If
    
    LogInformation Log, logString
    
    '------------------------- Try to get current version's number from Confluence ------------
    Dim strVersion As String
    Dim strMacDocs As String
    Dim strStyleDir As String
    Dim strFullVersionPath As String
    
    'Debug.Print FileName
    'Debug.Print InStrRev(FileName, ".do")
    strVersion = Left(FileName, InStrRev(FileName, ".do") - 1)
    strVersion = strVersion & ".txt"
    
    ' Always download version file to Style Directory - on Mac can't write to Startup w/o admin priv
    strStyleDir = StyleDir()
    
    strFullVersionPath = strStyleDir & Application.PathSeparator & strVersion
    'Debug.Print strVersion
    
    'If False, error in download; user was notified in DownloadFromConfluence function
    If DownloadFromConfluence(DownloadSource:=DownloadURL, FinalDir:=strStyleDir, LogFile:=Log, _
        FileName:=strVersion) = False Then
            NeedUpdate = False
            Exit Function
    End If
        
    '-------------------- Get version number of current template ---------------------
    If IsItThere(strFullVersionPath) = True Then
        NeedUpdate = True
        Dim strCurrentVersion As String

        strCurrentVersion = ReadTextFile(Path:=strFullVersionPath, FirstLineOnly:=True)
        
        ' git converts all line endings to LF which messes up PC, and I don't want to deal
        ' with it so we'll just remove everything
        If InStr(strCurrentVersion, vbCrLf) > 0 Then
            strCurrentVersion = Replace(strCurrentVersion, vbCrLf, "")
        ElseIf InStr(strCurrentVersion, vbLf) > 0 Then
            strCurrentVersion = Replace(strCurrentVersion, vbLf, "")
        ElseIf InStr(strCurrentVersion, vbCr) > 0 Then
            strCurrentVersion = Replace(strCurrentVersion, vbCr, "")
        End If
        
'        Debug.Print "Text File: |" & strCurrentVersion & "|"

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

Private Sub OpenDocMac(FilePath As String)
        Documents.Open FileName:=FilePath, ReadOnly:=True ', Visible:=False      'Mac Word 2011 doesn't allow Visible as an argument :(
End Sub

Private Sub OpenDocPC(FilePath As String)
        Documents.Open FileName:=FilePath, ReadOnly:=True, Visible:=False      'Win Word DOES allow Visible as an argument :)
End Sub



Private Function ImportVariable(strFile As String) As String
 
    Open strFile For Input As #1
    Line Input #1, ImportVariable
    Close #1
 
End Function


Private Sub CloseOpenDocs()

    '-------------Check for/close open documents---------------------------------------------
    Dim strInstallerName As String
    Dim strSaveWarning As String
    Dim objDocument As Document
    Dim B As Long
    Dim doc As Document
    
    strInstallerName = ThisDocument.Name

        'MsgBox "Installer Name: " & strInstallerName
        'MsgBox "Open docs: " & Documents.Count


    If Documents.Count > 1 Then
        strSaveWarning = "All other Word documents must be closed to run the macro." & vbNewLine & vbNewLine & _
            "Click OK and I will save and close your documents." & vbNewLine & _
            "Click Cancel to exit without running the macro and close the documents yourself."
        If MsgBox(strSaveWarning, vbOKCancel, "Close documents?") = vbCancel Then
            activeDoc.Close
            Exit Sub
        Else
            For Each doc In Documents
                On Error Resume Next        'To skip error if user is prompted to save new doc and clicks Cancel
                    'Debug.Print doc.Name
                    If doc.Name <> strInstallerName Then       'But don't close THIS document
                        doc.Save   'separate step to trigger Save As prompt for previously unsaved docs
                        doc.Close
                    End If
                On Error GoTo 0
            Next doc
        End If
    End If
    
End Sub



Private Function GetTemplatesList(TemplatesYouWant As TemplatesList, Optional PathToRepo As String) As Variant
    ' returns an array of paths to template files in their final installation locations
    ' if you want to use "allTemplates" (i.e., for updating code in templates), must include PathToRepo
    
    Dim strStartupDir As String
    Dim strStyleDir As String
    
    strStartupDir = Application.StartupPath
    strStyleDir = SharedFileInstaller.StyleDir()

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
        strPathsToTemplates(K) = strStyleDir & Application.PathSeparator & "Word-template.dotm"
    End If
    
    ' get the styles files for these requests
    If TemplatesYouWant = stylesTemplates Or _
        TemplatesYouWant = installTemplates Or _
        TemplatesYouWant = allTemplates Then
        K = K + 1
        ReDim Preserve strPathsToTemplates(1 To K)
        strPathsToTemplates(K) = strStyleDir & Application.PathSeparator & "macmillan.dotx"
        
        K = K + 1
        ReDim Preserve strPathsToTemplates(1 To K)
        strPathsToTemplates(K) = strStyleDir & Application.PathSeparator & "macmillan_NoColor.dotx"

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

