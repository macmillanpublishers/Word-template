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

