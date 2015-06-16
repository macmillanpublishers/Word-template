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

'======== PURPOSE ===================================
'Downloads and installs a single template file in the Startup directory on a PC & logs the download

'======== USE =======================================
'This is Part 1 of 2. Part 2 declares global variables in a module named "GlobalVars"
'That is where you'll enter your template file name
'The template file needs to be uploaded as an attachment to https://confluence.macmillan.com/display/PBL/Test
'This Part 1 code needs to reside in the ThisDocument module of a .docm file so that
'it will launch when users open the file.

    Call DeclareGlobalVars

    Dim TheOS As String
    TheOS = System.OperatingSystem
    
    'Doesn't work on Mac
    If TheOS Like "*Mac*" Then
        MsgBox "This installer works for PC only. To install the " & strTemplateName & _
            "on a Mac, please install from Self Service, Digital Workflow category."
            ActiveDocument.Close (wdDoNotSaveChanges)
            Exit Sub
    Else
        'Alert user to install
        Dim strWelcome As String
    
        strWelcome = "Welcome to the " & strTemplateName & " Installer!" & vbNewLine & vbNewLine & _
            "Please click OK to begin the installation. It should only take a few seconds."
    
        If MsgBox(strWelcome, vbOKCancel, strTemplateName) = vbCancel Then
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
        
        '-----------------Define variables--------------------------------------------------
        'For template file
        Dim strStartupDir As String
        Dim strGtFinalPath As String
        
        'For style templates & log file
        Dim strStyleDir As String
        
        'For log files
        Dim strLogDir As String
        Dim strLogFile As String
        Dim strLogPath As String
        Dim logString As String

        strStartupDir = Application.StartupPath
        strGtFinalPath = strStartupDir & "\" & strTemplateFile
        
        strStyleDir = Environ("PROGRAMDATA") & "\MacmillanStyleTemplate"
        
        strLogDir = strStyleDir & "\log"
        strLogFile = "mac_templates_" & Format(Date, "yyyy-mm-dd") & "_" & Format(Time, "hh-mm-ss") & ".log"
        strLogPath = strLogDir & "\" & strLogFile
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
                    "DOWNLOAD " & strTemplateFile & vbNewLine & _
                    "------------------------------------------"
        LogInformation strLogPath, logString
        
        'If False, error in download; user was notified in DownloadFromConfluence function
        If DownloadFromConfluence(strGtFinalPath, strLogPath) = False Then
            ActiveDocument.Quit (wdDoNotSaveChanges)
            Exit Sub
        End If
    End If
    
    '------Display installation complete message and close doc (ending sub)---------------
    Dim strComplete As String
    
    strComplete = "The " & strTemplateName & " has been installed on your computer." & vbNewLine & vbNewLine & _
        "Click OK to close Word. When you restart Word, the template will be available."
        
    MsgBox strComplete, vbOKOnly, "Installation Successful"
    
    '------Close and restart Word for template changes to take effect---------------------
    'Would love to get restart to work...
    'Dim restartTime As Variant
    'restartTime = Now + TimeValue("00:00:01")
    'Application.OnTime When:=restartTime, Name:="Restart"
    Application.Quit SaveChanges:=wdDoNotSaveChanges          'DEBUG: comment out this line
    
End Sub
Private Function DownloadFromConfluence(FinalPath As String, LogFile As String) As Boolean
        
    Dim logString As String
    Dim strTmpPath As String
    Dim strErrMsg As String
        
    logString = ""
    strTmpPath = Environ("TEMP") & "\" & strTemplateFile
        
    'try to download the file from Public Confluence page
    Dim myURL As String
    Dim WinHttpReq As Object
    Dim oStream As Object
    
    'this is download link, actual page housing files is https://confluence.macmillan.com/display/PBL/Test
    myURL = "https://confluence.macmillan.com/download/attachments/9044274/" & strTemplateFile
        
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
                MsgBox strErrMsg, vbCritical, "Error 1: Connection error (" & strTemplateFile & ")"
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
        MsgBox strErrMsg, vbCritical, "Error 2: Http status " & WinHttpReq.Status & " (" & strTemplateFile & ")"
        DownloadFromConfluence = False
        Exit Function
    End If
        
    'Error if download was not successful
    If Dir(strTmpPath) = vbNullString Then
        logString = "-- " & strTemplateFile & " file download to Temp was not successful. Exiting installer."
        LogInformation LogFile, logString
        strErrMsg = "There was an error downloading the Macmillan template." & vbNewLine & _
            "Please contact workflows@macmillan.com for assitance."
        MsgBox strErrMsg, vbCritical, "Error 3: Download failed (" & strTemplateFile & ")"
        DownloadFromConfluence = False
        Exit Function
    Else
        logString = "-- " & strTemplateFile & " file download to Temp was successful."
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
                logString = "-- old " & strTemplateFile & " file is open, can't delete/replace. Alerting user, exiting sub."
                LogInformation LogFile, logString
                strErrMsg = "Please close all other Word documents and try again."
                MsgBox strErrMsg, vbCritical, "Error 4: Previous version removal failed (" & strTemplateFile & ")"
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
        logString = "-- Final directory clear of " & strTemplateFile & " file."
        LogInformation LogFile, logString
        Name strTmpPath As FinalPath
    Else
        logString = "-- old " & strTemplateFile & " file not cleared from Final directory. Exiting installer."
        LogInformation LogFile, logString
        strErrMsg = "There was an error installing the Macmillan template." & vbNewLine & _
            "Please close all other Word documents and try again, or contact workflows@macmillan.com."
        MsgBox strErrMsg, vbCritical, "Error 5: Previous version uninstall failed (" & strTemplateFile & ")"
        DownloadFromConfluence = False
        Exit Function
    End If
    
    'If move was successful, yay! Else, :(
    If Dir(FinalPath) <> vbNullString Then
        logString = "-- " & strTemplateFile & " file successfully saved to final directory."
        LogInformation LogFile, logString
    Else
        logString = "-- " & strTemplateFile & " file not saved to final directory."
        LogInformation LogFile, logString
        strErrMsg = "There was an error installing the Macmillan template." & vbNewLine & vbNewLine & _
            "Please cotact workflows@macmillan.com for assistance."
        MsgBox strErrMsg, vbCritical, "Error 6: Installation failed (" & strTemplateFile & ")"
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

Sub Restart()
    Documents.Add
    MsgBox "The installation has been successful."
End Sub


