VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Document_Open()
'created by Erica Warren - erica.warren@macmillan.com
'installs the MacmillanGT.dotm template in STARTUP dir
    
    Dim TheOS As String
    TheOS = System.OperatingSystem
    
    'Doesn't work on Mac
    If TheOS Like "*Mac*" Then
        MsgBox "This installer works for PC only. To install the Macmillan Style Templates " & _
            "on a Mac, please do the following:" & vbNewLine & vbNewLine & _
            "In-house: Install from Self Service, Digital Workflow category." & vbNewLine & _
            "External: Follow the instructions at confluence.macmillan.com/display/PBL/Install+the+Macmillan+Template."
            ActiveDocument.Close (wdDoNotSaveChanges)
    Else
        'Alert user to install
        Dim strWelcome As String
    
        strWelcome = "Welcome to the Macmillan Style Template Installer!" & vbNewLine & vbNewLine & _
            "Please click OK to begin the installation. It should only take a few seconds."
    
        If MsgBox(strWelcome, vbOKCancel, "Macmillan Style Template") = vbCancel Then
            ActiveDocument.Close (wdDoNotSaveChanges)
            Exit Sub
        End If
        
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
                "Click OK and I will save and close your documents." & vbNewLine & vbNewLine & _
                "Click Cancel to exit without installing and close the documents yourself."
            If MsgBox(strSaveWarning, vbOKCancel, "Close documents?") = vbCancel Then
                ActiveDocument.Close
                Exit Sub
            Else
                For b = 1 To Documents.Count
                    'Debug.Print "Current doc " & b & ": " & Documents(b).Name
                    On Error Resume Next        'To trap error if user is prompted to save new doc and clicks Cancel
                        If Documents(b).Name <> strInstallerName Then
                            Documents(b).Save   'separate step to trigger Save As prompt for previously unsaved docs
                            Documents(b).Close
                        End If
                    On Error GoTo 0
                Next b
            End If
        End If
        
        '-----------------Define variables--------------------------------------------------
        'For MacmillanGT.dotm
        Dim strGtFile As String
        Dim strStartupDir As String
        Dim strGtFinalPath As String
        
        'For style templates
        Dim strStyleDir As String
        
        'For log files
        Dim strLogDir As String
        Dim strLogFile As String
        Dim strLogPath As String
        Dim logString As String

        strGtFile = "MacmillanGT.dotm"
        strStartupDir = Application.StartupPath
            'Debug.Print "STARTUP: " & strStartupDir
        strGtFinalPath = strStartupDir & "\" & strGtFile
            'Debug.Print "Full STARTUP path: " & strGtFinalPath
        
        strStyleDir = Environ("PROGRAMDATA") & "\MacmillanStyleTemplate"
            'Debug.Print "Style Template path: " & strStyleDir
        
        strLogDir = strStyleDir & "\log"
        strLogFile = "mac_templates_" & Format(Date, "yyyy-mm-dd") & "_" & Format(Time, "hh-mm-ss") & ".log"
            'Debug.Print strLogFile
        strLogPath = strLogDir & "\" & strLogFile
            'Debug.Print "Log file path: " & strLogPath
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
        
        '-------------download MacmillanGT file------------------------------------------
        'Log attempt to download
        logString = "------------------------------------------" & vbNewLine & _
                    "DOWNLOAD " & strGtFile & vbNewLine & _
                    "------------------------------------------"
        LogInformation strLogPath, logString
        
        'Download GT file
        If DownloadFromConfluence(strGtFile, strGtFinalPath, strLogPath) = False Then
            ActiveDocument.Close (wdDoNotSaveChanges)
            Exit Sub
        End If
        
        '------------------------download style templates--------------------------------
        
        Dim arrStyleTemplates() As String
        ReDim arrStyleTemplates(1 To 3)
        Dim strTemplateFinalPath As String
        Dim a As Long
        
        arrStyleTemplates(1) = "macmillan.dotm"
        arrStyleTemplates(2) = "macmillan_NoColor.dotm"
        arrStyleTemplates(3) = "MacmillanCoverCopy.dotm"
        
        For a = 1 To UBound(arrStyleTemplates())
            'Log attempt to download
            logString = "------------------------------------------" & vbNewLine & _
                        "DOWNLOAD " & arrStyleTemplates(a) & vbNewLine & _
                        "------------------------------------------"
            LogInformation strLogPath, logString
            
            strTemplateFinalPath = strStyleDir & "\" & arrStyleTemplates(a)
            
            If DownloadFromConfluence(arrStyleTemplates(a), strTemplateFinalPath, strLogPath) = False Then
                ActiveDocument.Close (wdDoNotSaveChanges)       'If ANY templates do not download correctly, sub will end.
                Exit Sub
            End If
        Next a
            
    End If
    
    '------Display installation complete message and close doc (ending sub)---------------
    Dim strComplete As String
    strComplete = "The Macmillan templates have been installed on your computer." & vbNewLine & vbNewLine & _
        "Close all Word files, then open Word again for the new templates to take effect."
    MsgBox strComplete, vbOKOnly, "Installation Successful"
    ActiveDocument.Close (wdDoNotSaveChanges)          'DEBUG: comment out this line
    
    
End Sub
Private Function DownloadFromConfluence(FileName As String, FinalPath As String, LogFile As String) As Boolean
        
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
    If InStr(LCase(FinalPath), LCase("startup")) > 0 Then
        On Error Resume Next            'Error = add-in not available, don't need to uninstall
            AddIns(FinalPath).Installed = False
        On Error GoTo 0
    End If
    
    'If file exists already, log it and delete it
    If Dir(FinalPath) <> vbNullString Then
        logString = "-- Previous version file in final directory."
        LogInformation LogFile, logString
        
        On Error Resume Next
            Kill FinalPath
        
            'If an error occurs, the file is currently open.
            If Err.Number <> 0 Then 'Add actual error number?
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


