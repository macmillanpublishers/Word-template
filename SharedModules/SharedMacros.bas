Attribute VB_Name = "SharedMacros"
' For macros that are shared by macros in other modules of the Macmillan template
' All should be declared as Public for use from other modules

Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub Doze(ByVal lngPeriod As Long)
    DoEvents
    Sleep lngPeriod
    ' Call it in desired location to sleep for 1 seconds like this:
    ' Doze 1000
End Sub
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
        'Debug.Print "IsItThere Error " & Err.Number & ": " & Err.Description
    End If
End Function

Public Function DownloadFromConfluence(StagingURL As Boolean, FinalDir As String, LogFile As String, FileName As String) As Boolean
'FinalPath is directory w/o file name

    Dim logString As String
    Dim strTmpPath As String
    Dim strBashTmp As String
    Dim strMacHD As String
    Dim strFinalPath As String
    Dim strErrMsg As String
    Dim myURL As String
    
    strFinalPath = FinalDir & Application.PathSeparator & FileName
    
    'Get URL to download from
    If StagingURL = True Then
        'actual page to update files is https://confluence.macmillan.com/display/PBL/Word+template+downloads+-+staging
        myURL = "https://confluence.macmillan.com/download/attachments/35001370/" & FileName
    Else
        'actual page to update files is https://confluence.macmillan.com/display/PBL/Word+template+downloads+-+production
        myURL = "https://confluence.macmillan.com/download/attachments/9044274/" & FileName
    End If
    
    'Get temp dir based on OS, then download file.
    #If Mac Then
        'set tmp dir
        strMacHD = "Macintosh HD"
        strTmpPath = strMacHD & ":private:tmp" & Application.PathSeparator & FileName
        strBashTmp = Replace(Right(strTmpPath, Len(strTmpPath) - Len(strMacHD)), ":", "/")
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
            'But first make sure file isn't already there, delete if it is
            If IsItThere(strTmpPath) = True Then
                Kill strTmpPath
            End If
            
            'Now download the file
            ShellAndWaitMac ("rm -f " & strBashTmp & " ; curl -o " & strBashTmp & " " & myURL)
            'ShellAndWaitMac ("rm -f /private/tmp/MacmillanGT.dotm ; curl -o /private/tmp/MacmillanGT.dotm " & myURL)
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
        
    'If delete was successful, move downloaded file to final directory
    If IsItThere(strFinalPath) = False Then
        logString = Now & " -- Final directory clear of " & FileName & " file."
        LogInformation LogFile, logString
        
        'Mac won't load macros from a template downloaded from the internet to Startup.
        'Need to open and save as to final location for macros to work.
        #If Mac Then
            If InStr(1, FileName, ".dotm") > 0 And InStr(1, LCase(strFinalPath), LCase("startup"), vbTextCompare) > 0 Then      'File is a template being saved in startup dir
                Documents.Open FileName:=strTmpPath, ReadOnly:=True ', Visible:=False doesn't work on Mac
                Documents(strTmpPath).SaveAs (strFinalPath)
                Documents(strFinalPath).Close
            Else
                Name strTmpPath As strFinalPath
            End If
        #Else
            Name strTmpPath As strFinalPath
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
    
    DownloadFromConfluence = True

End Function
 
Public Function ShellAndWaitMac(cmd As String) As String

    Dim result As String
    Dim scriptCmd As String ' Macscript command
    
    scriptCmd = "do shell script """ & cmd & """"
    result = MacScript(scriptCmd) ' result contains stdout, should you care
    ShellAndWaitMac = result

End Function

Public Sub LogInformation(LogFile As String, LogMessage As String)

Dim FileNum As Integer
    FileNum = FreeFile ' next file number
    Open LogFile For Append As #FileNum ' creates the file if it doesn't exist
    Print #FileNum, LogMessage ' write information at the end of the text file
    Close #FileNum ' close the file
End Sub

Public Function CreateLogFileInfo(ByRef FileName As String) As Variant
' Creates the style dir, log dir, and log file name variables for use in other subs.
' File name should not contain periods other than before file type

    Dim strLogFile As String
    Dim strUser As String
    Dim strStyle As String
    Dim strLogFolder As String
    Dim strLogPath As String
    
    'Create logfile name
    strLogFile = Left(FileName, InStrRev(FileName, ".") - 1)
    strLogFile = strLogFile & "_updates.log"
    
    'Create directory names based on OS
    #If Mac Then
        Dim strUser As String
        strUser = MacScript("tell application " & Chr(34) & "System Events" & Chr(34) & Chr(13) & _
            "return (name of current user)" & Chr(13) & "end tell")
        strStyle = "Macintosh HD:Users:" & strUser & ":Documents:MacmillanStyleTemplate"
        strLogFolder = strStyleDir & Application.PathSeparator & "log"
        strLogPath = strLogDir & Application.PathSeparator & strLogFile
    #Else
        strStyle = Environ("ProgramData") & "\MacmillanStyleTemplate"
        strLogFolder = strStyle & Application.PathSeparator & "log"
        strLogPath = strLogFolder & Application.PathSeparator & strLogFile
    #End If
    'Debug.Print strLogPath

    Dim arrFinalDirs() As Variant
    ReDim arrFinalDirs(1 To 3)
    
    arrFinalDirs(1) = strStyle
    arrFinalDirs(2) = strLogFolder
    arrFinalDirs(3) = strLogPath
    
    CreateLogFileInfo = arrFinalDirs

End Function

Public Function CheckLog(StyleDir As String, LogDir As String, LogPath As String) As Boolean
'LogPath is *full* path to log file, including file name. Created by CreateLogFileInfo sub, to be called before this one.

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
