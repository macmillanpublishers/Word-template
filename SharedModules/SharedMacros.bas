Attribute VB_Name = "SharedMacros"
Option Explicit
' For macros that are shared by macros in other modules of the Macmillan template

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
    
Exit Function

ErrHandler:
    If Err.Number = 68 Then     ' "Device unavailable"
        IsItThere = False
    Else
        'Debug.Print "IsItThere Error " & Err.Number & ": " & Err.Description
    End If
End Function

Private Function DownloadFromConfluence(FinalDir As String, LogFile As String, FileName As String) As Boolean
'FinalPath is directory w/o file name

    Dim logString As String
    Dim strTmpPath As String
    Dim strBashTmp As String
    Dim strMacHD As String
    Dim strFinalPath As String
    Dim strErrMsg As String
    Dim myURL As String
    
    strFinalPath = FinalDir & Application.PathSeparator & FileName
    
    'this is download link, actual page housing files is https://confluence.macmillan.com/display/PBL/Test
    myURL = "https://confluence.macmillan.com/download/attachments/9044274/" & FileName
            
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
            ShellAndWaitMac ("rm -f " & strBashTmp & " ; curl -o " & strBashTmp & " " & myURL)
            'ShellAndWaitMac ("rm -f /private/tmp/MacmillanGT.dotm ; curl -o /private/tmp/MacmillanGT.dotm " & myURL)
        End If
    #Else
        'set tmp dir
        strTmpPath = Environ("TEMP") & Application.PathSeparator & FileName 'Environ gives temp dir for Mac too?
    
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
        
    'If delete was successful, move downloaded file to final directory
    If IsItThere(strFinalPath) = False Then
        logString = Now & " -- Final directory clear of " & FileName & " file."
        LogInformation LogFile, logString
        
        'Mac won't load macros from a template downloaded from the internet to Startup. Need to open and save as to final location for macros to work.
        #If Mac Then
            If InStr(1, FileName, ".dotm") > 0 And InStr(1, LCase(strFinalPath), LCase("startup"), vbTextCompare) > 0 Then      'File is a template being saved in startup dir
                Documents.Open FileName:=strTmpPath, ReadOnly:=True ', Visible:=False
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


