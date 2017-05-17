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

Public Sub Installer(Installer As Boolean, TemplateName As String, ByRef _
  TemplatesToInstall As Collection)

'"Installer" argument = True if this is for a standalone installtion file.
'"Installer" argument = False is this is part of a daily check of the current
' file and only updates if out of date.

' Remove file names from collection if they don't need to be updated rignt now
  Dim strFileName As Variant
  Dim dictFileInfo As Dictionary

  If Installer = False Then
    For Each strFileName In TemplateToInstall
    ' If exists, check if it's been checked today
      If IsTemplateThere(FileName:=strFileName) = True Then
        Set dictFileInfo = FileInfo(strFileName)
      ' If HASN'T been checked today, run NeedUpdate
        If CheckLog(LogPath:=dictFileInfo("Log")) = False Then
        ' If DON'T need update, remove from collection
          If NeedUpdate(FileName:=strFileName) = False Then
            TemplatesToInstall.Remove strFileName
          End If
        Else ' Has been checked today, don't update
          TemplatesToInstall.Remove strFileName
        End If
      End If
    Next strFileName
  End If

' If everything is OK, quit sub
  If TemplatesToInstall.Count = 0 Then
    Exit Sub
  End If

  ' Alert user that installation is happening
  Dim strWelcome As String
  If Installer = True Then
    strWelcome = "Welcome to the " & TemplateName & " Installer!" & vbNewLine _
    & vbNewLine & "Please click OK to begin the installation. It should only" _
    & " take a few seconds."
  Else
    strWelcome = "Your " & TemplateName & " is out of date. Click OK to " & _
      "update automatically."
  End If
  
  If MsgBox(strWelcome, vbOKCancel, TemplateName) = vbCancel Then
    MsgBox "Please try to install the files at a later time."
  
    If Installer = True Then
      activeDoc.Close (wdDoNotSaveChanges)
    End If
    Exit Sub
  End If

' ---------------- Close any open docs (with prompt) --------------------------
  Call CloseOpenDocs

' --------- DOWNLOAD FILES ----------------------------------------------------
'If False, error in download; user was notified in DownloadFromGithub function
  For Each strFileName In TemplatesToInstall
    If DownloadFromGithub(FileName:=strFileName) = False Then
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

  ' If we just updated the main template, delete the old toolbar
  ' Will be added again by Word-template AutoExec when it's launched
    #If Mac Then
      Dim Bar As CommandBar
      If strFileName = "Word-template.dotm" Then
        For Each Bar In CommandBars
          If Bar.Name = "Macmillan Tools" Then
            Bar.Delete
          End If
        Next
      End If
    #End If
  Next strFileName
  
'------Display installation complete message   ---------------------------
  Dim strComplete As String
  Dim strInstallType As String
  
' Quit if it's an installer, but not if it's an updater
' (updater was causing conflicts between Word-template.dotm and GtUpdater)
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

' Create directory if it doesn't exist yet
  If Utils.IsItThere(strFullPath) = False Then
    MkDir strFullPath
  End If

  StyleDir = strFullPath

End Function

' ===== DownloadJson ==========================================================
' Downloads JSON, if download fails can still continue if a previous version
' is available.
' PARAMS
' FileName: name and extension of JSON file (not path)

' RETURNS:
' Full path to downloaded JSON file.
' If download failed and we don't already have a copy locally, returns null string

Public Function DownloadJson(FileName As String) As String
  Dim dictJsonInfo As Dictionary
  Dim strFinalPath As String
  
  Set dictJsonInfo = FileInfo(FileName)
  strFinalPath = dictJsonInfo("Final")

  If DownloadFromGithub(FileName) = False Then
    If Utils.IsItThere(strFinalPath) = True Then
      DownloadJson = strFinalPath
    Else
      DownloadJson = vbNullString
    End If
  Else
    DownloadJson = strFinalPath
  End If

End Function


Public Function DownloadCSV(FileName As String) As Variant
'---------Download CSV with design specs from Confluence site-------
  Dim dictCsvInfo As Dictionary
  Set dictCsvInfo = FileInfo(FileName)

' Always download, so don't return CheckLog resutl
  CheckLog LogPath:=dictCsvInfo("Log")

  If DownloadFromGithub(FileName:=FileName) = False Then
    ' If download fails, check if we have an older version of the CSV
    If IsItThere(dictCsvInfo("Final")) = False Then
      strMessage = "Looks like we can't download the design info from the " & _
        "internet right now. Please check your internet connection, or " & _
        "contact workflows@macmillan.com."
      MsgBox strMessage, vbCritical, "Error 5: Download failed, no design file."
      Exit Function
    Else
      strMessage = "Looks like we can't download the most up-to-date design " _
        & "info from the internet right now, so we'll just use the info we " & _
        "have on file for your castoff."
      MsgBox strMessage, vbInformation, "Let's do this thing!"
    End If
  End If
    
' Heading row/col different based on different InfoTypes
  Dim blnRemoveHeaderRow As Boolean
  Dim blnRemoveHeaderCol As Boolean
  
' Because castoff CSV has header row and col, but Spine CSV only has header row
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
    strMessage = "The macro is unable to access the data file right now." _
      & " Please check your internet connection and try again, or " & _
      "contact workflows@macmillan.com."
    MsgBox strMessage, vbCritical, "Error 3: CSV doesn't exist"
    Exit Function
  Else
  ' Load CSV into an array
    arrFinal = LoadCSVtoArray(Path:=strPath, RemoveHeaderRow:= _
      blnRemoveHeaderRow, RemoveHeaderCol:=blnRemoveHeaderCol)
  End If

  DownloadCSV = arrFinal

End Function


Public Function GetTemplatesList(TemplatesYouWant As TemplatesList) As _
  Collection
' returns a Collection of template file names to download

  Dim colTemplates As Collection
  Set colTemplates = New Collection

' get the updater file for these requests
  If TemplatesYouWant = updaterTemplates Or _
    TemplatesYouWant = installTemplates Or _
    TemplatesYouWant = allTemplates Then
      colTemplates.Add "GtUpdater.dotm"
  End If

' get the tools file for these requests
  If TemplatesYouWant = toolsTemplates Or _
    TemplatesYouWant = installTemplates Or _
    TemplatesYouWant = allTemplates Then
      colTemplates.Add "Word-template.dotm"
  End If

  ' get the styles files for these requests
  If TemplatesYouWant = stylesTemplates Or _
    TemplatesYouWant = installTemplates Or _
    TemplatesYouWant = allTemplates Then
    colTemplates.Add "macmillan.dotx"
    colTemplates.Add "macmillan_NoColor.dotx"
    colTemplates.Add "macmillan_CoverCopy.dotm"
  End If

  ' also get the installer file
  If TemplatesYouWant = allTemplates And PathToRepo <> vbNullString Then
    colTemplates.Add "MacmillanTemplateInstaller.docm"
  End If

  Set GetTemplatesList = colTemplates

End Function


' ===== DownloadFromGithub ================================================
' Actually now it downloads from Github but don't want to mess with things, we're
' going to be totally refacroting soon.

' DEPENDENCIES:
' Add file and download URL info to FullURL function.

Private Function DownloadFromGithub(FileName As String) As Boolean

  Dim dictFullPaths As Dictionary
  Set dictFullPaths = FileInfo(FileName)
  
  Dim strErrMsg As String
  Dim myURL As String
  Dim logString As String

' Get URL to download from.
  myURL = FullURL(FileName:=FileName)

  'Get temp dir based on OS, then download file.
  #If Mac Then
    Dim strBashTmp As String
    strBashTmp = Replace(Right(dictFullPaths("Tmp"), Len(dictFullPaths("Tmp")) - (InStr(dictFullPaths("Tmp"), ":") - 1)), ":", "/")
    'DebugPrint strBashTmp
    
    'check for network
    If ShellAndWaitMac("ping -o google.com &> /dev/null ; echo $?") <> 0 Then   'can't connect to internet
      logString = Now & " -- Tried update; unable to connect to network."
      LogInformation dictFullPaths("Log"), logString
      strErrMsg = "There was an error trying to download the Macmillan template." & vbNewLine & vbNewLine & _
                  "Please check your internet connection or contact workflows@macmillan.com for help."
      MsgBox strErrMsg, vbCritical, "Error 1: Connection error (" & FileName & ")"
      DownloadFromGithub = False
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
        LogInformation dictFullPaths("Log"), logString
        strErrMsg = "It looks like that file isn't available for download." & vbNewLine & vbNewLine & _
                    "Please contact workflows@macmillan.com for help."
        MsgBox strErrMsg, vbCritical, "Error 7: File not found (" & FileName & ")"
        DownloadFromGithub = False
        Exit Function
      Else
        logString = Now & " -- Http status is " & httpStatus & ". Cannot download file."
        LogInformation dictFullPaths("Log"), logString
        strErrMsg = "There was an error trying to download the Macmillan templates." & vbNewLine & vbNewLine & _
            "Please check your internet connection or contact workflows@macmillan.com for help."
        MsgBox strErrMsg, vbCritical, "Error 2: Http status " & httpStatus & " (" & FileName & ")"
        DownloadFromGithub = False
        Exit Function
      End If
    End If
  #Else
    'Check if file is already in tmp dir, delete if yes
    If IsItThere(dictFullPaths("Tmp")) = True Then
      Kill dictFullPaths("Tmp")
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
        'DebugPrint WinHttpReq.Status
        logString = Now & " -- could not connect to Confluence site: Error " & Err.Number
        LogInformation dictFullPaths("Log"), logString
        strErrMsg = "There was an error trying to download the Macmillan template." & vbNewLine & vbNewLine & _
            "Please check your internet connection or contact workflows@macmillan.com for help."
        MsgBox strErrMsg, vbCritical, "Error 1: Connection error (" & FileName & ")"
        DownloadFromGithub = False
        On Error GoTo 0
        Exit Function
      End If
    On Error GoTo 0
        
    'DebugPrint "Http status for " & FileName & ": " & WinHttpReq.Status
    If WinHttpReq.Status = 200 Then  ' 200 = HTTP request is OK
  
      'if connection OK, download file to temp dir
      myURL = WinHttpReq.responseBody
      Set oStream = CreateObject("ADODB.Stream")
      oStream.Open
      oStream.Type = 1
      oStream.Write WinHttpReq.responseBody
      oStream.SaveToFile dictFullPaths("Tmp"), 2 ' 1 = no overwrite, 2 = overwrite
      oStream.Close
      Set oStream = Nothing
      Set WinHttpReq = Nothing
    ElseIf WinHttpReq.Status = 404 Then ' 404 = file not found
      logString = Now & " -- 404 File not found. Cannot download file."
      LogInformation dictFullPaths("Log"), logString
      strErrMsg = "It looks like that file isn't available for download." & vbNewLine & vbNewLine & _
          "Please contact workflows@macmillan.com for help."
      MsgBox strErrMsg, vbCritical, "Error 7: File not found (" & FileName & ")"
      DownloadFromGithub = False
      Exit Function
    Else
      logString = Now & " -- Http status is " & WinHttpReq.Status & ". Cannot download file."
      LogInformation dictFullPaths("Log"), logString
      strErrMsg = "There was an error trying to download the Macmillan templates." & vbNewLine & vbNewLine & _
          "Please check your internet connection or contact workflows@macmillan.com for help."
      MsgBox strErrMsg, vbCritical, "Error 2: Http status " & WinHttpReq.Status & " (" & FileName & ")"
      DownloadFromGithub = False
      Exit Function
    End If
  #End If
        
  'Error if download was not successful
  If IsItThere(dictFullPaths("Tmp")) = False Then
    logString = Now & " -- " & FileName & " file download to Temp was not successful."
    LogInformation dictFullPaths("Log"), logString
    strErrMsg = "There was an error downloading the Macmillan template." & vbNewLine & _
        "Please contact workflows@macmillan.com for assitance."
    MsgBox strErrMsg, vbCritical, "Error 3: Download failed (" & FileName & ")"
    DownloadFromGithub = False
    On Error GoTo 0
    Exit Function
  Else
    logString = Now & " -- " & FileName & " file download to Temp was successful."
    LogInformation dictFullPaths("Log"), logString
  End If

  'If file exists already, log it and delete it
  If IsItThere(dictFullPaths("Final")) = True Then

    logString = Now & " -- Previous version file in final directory."
    LogInformation dictFullPaths("Log"), logString
    
    ' get file extension
    Dim strExt As String
    strExt = Utils.GetFileExtension(dictFullPaths("Final"))
    
    ' can't delete template if it's installed as an add-in
    If InStr(strExt, "dot") > 0 Then
      Utils.UnloadAddIn (dictFullPaths("Final"))
    End If

    ' Test if dir is read only
    Dim strFinalDir As String
    strFinalDir = Utils.GetPath(dictFullPaths("Final"))
    If IsReadOnly(strFinalDir) = True Then ' Dir is read only
      logString = Now & " -- old " & FileName & " file is read only, can't delete/replace. " _
          & "Alerting user."
      LogInformation dictFullPaths("Log"), logString
      strErrMsg = "The installer doesn't have permission. Please conatct workflows" & _
          "@macmillan.com for help."
      MsgBox strErrMsg, vbCritical, "Error 8: Permission denied (" & FileName & ")"
      DownloadFromGithub = False
      On Error GoTo 0
      Exit Function
    Else
      On Error Resume Next
        Kill dictFullPaths("Final")
        
        If Err.Number = 70 Then         'File is open and can't be replaced
          logString = Now & " -- old " & FileName & " file is open, can't delete/replace. Alerting user."
          LogInformation dictFullPaths("Log"), logString
          strErrMsg = "Please close all other Word documents and try again."
          MsgBox strErrMsg, vbCritical, "Error 4: Previous version removal failed (" & FileName & ")"
          DownloadFromGithub = False
          On Error GoTo 0
          Exit Function
        End If
      On Error GoTo 0
    End If
  Else
    logString = Now & " -- No previous version file in final directory."
    LogInformation dictFullPaths("Log"), logString
  End If
      
  'If delete was successful, move downloaded file to final directory
  If IsItThere(dictFullPaths("Final")) = False Then
    logString = Now & " -- Final directory clear of " & FileName & " file."
    LogInformation dictFullPaths("Log"), logString
    
    ' move template to final directory
    Name dictFullPaths("Tmp") As dictFullPaths("Final")
    
    'Mac won't load macros from a template downloaded from the internet to Startup.
    'Need to send these commands for it to work, see Confluence
    ' Do NOT use open/save as option, this removes customUI which creates Mac Tools toolbar later
    #If Mac Then
      If InStr(1, FileName, ".dotm") Then
        Dim strCommand As String
        strCommand = "do shell script " & Chr(34) & "xattr -wx com.apple.FinderInfo \" & Chr(34) & _
            "57 58 54 4D 4D 53 57 44 00 10 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00\" & _
            Chr(34) & Chr(32) & Chr(34) & " & quoted form of POSIX path of " & Chr(34) & dictFullPaths("Final") & Chr(34)
            'DebugPrint strCommand
            MacScript (strCommand)
      End If
    #End If
  
  Else
    logString = Now & " -- old " & FileName & " file not cleared from Final directory."
    LogInformation dictFullPaths("Log"), logString
    strErrMsg = "There was an error installing the Macmillan template." & vbNewLine & _
        "Please close all other Word documents and try again, or contact workflows@macmillan.com."
    MsgBox strErrMsg, vbCritical, "Error 5: Previous version uninstall failed (" & FileName & ")"
    DownloadFromGithub = False
    On Error GoTo 0
    Exit Function
  End If
  
  'If move was successful, yay! Else, :(
  If IsItThere(dictFullPaths("Final")) = True Then
    logString = Now & " -- " & FileName & " file successfully saved to final directory."
    LogInformation dictFullPaths("Log"), logString
  Else
    logString = Now & " -- " & FileName & " file not saved to final directory."
    LogInformation dictFullPaths("Log"), logString
    strErrMsg = "There was an error installing the Macmillan template." & vbNewLine & vbNewLine & _
        "Please cotact workflows@macmillan.com for assistance."
    MsgBox strErrMsg, vbCritical, "Error 6: Installation failed (" & FileName & ")"
    DownloadFromGithub = False
    On Error GoTo 0
    Exit Function
  End If
  
  'Cleanup: Get rid of temp file if downloaded correctly
  If IsItThere(dictFullPaths("Tmp")) = True Then
    Kill dictFullPaths("Tmp")
  End If
  
  ' Disable Startup add-ins so they don't launch right away and mess of the code that's running
  If InStr(1, LCase(dictFullPaths("Final")), LCase("startup"), vbTextCompare) > 0 Then         'LCase because "startup" was staying in all caps for some reason, UCase wasn't working
    On Error Resume Next                                        'Error = add-in not available, don't need to uninstall
      AddIns(dictFullPaths("Final")).Installed = False
    On Error GoTo 0
  End If
  
  DownloadFromGithub = True

End Function


Private Sub LogInformation(LogFile As String, LogMessage As String)

' Create parent dir if it doesn't exist yet
  If Utils.ParentDirExists(LogFile) = False Then
    MkDir Utils.GetPath(LogFile)
  End If

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
        'DebugPrint Path
        
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
                    'DebugPrint one_line(c)
                    the_array((R - lngHeaderRow), (C - lngHeaderCol)) = one_line(C)   ' -1 because if are not using header row/column from CSV
                Next C
            End If
        Next R
    
        ' Prove we have the data loaded.
'         DebugPrint LBound(the_array)
'         DebugPrint UBound(the_array)
'         For R = 0 To (num_rows - 1)          ' -1 again if we removed the header row
'             For c = 0 To num_cols      ' -1 again if we removed the header column
'                 DebugPrint the_array(R, c) & " | ";
'             Next c
'             DebugPrint
'         Next R
'         DebugPrint "======="
    
    LoadCSVtoArray = the_array
 
End Function



Private Function CheckLog(LogPath As String) As Boolean
' LogPath = full path to log file we're checking
' REturns TRUE if file has been updated today.
  Dim logString As String
  Dim strLogDir As String
  Dim strStylesDir As String
  
  strLogDir = Utils.GetPath(LogPath)
  strStylesDir = StyleDir

' Have to create "log" directory here, bc creation elsewhere marks as "updated"
  If IsItThere(LogPath) = False Then
    CheckLog = False
    logString = Now & " -- Creating logfile."
    If IsItThere(strLogDir) = False Then
      If IsItThere(strStylesDir) = False Then
        MkDir (strStylesDir)
        MkDir (strLogDir)
        logString = Now & " -- Creating MacmillanStyleTemplate directory."
      Else
        MkDir (strLogDir)
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


Private Function IsTemplateThere(FileName As String)
  Dim dictTemplateInfo As Dictionary
  Set dictTemplateInfo = FileInfo(FileName)

  Dim logString As String
  Dim strDir As String
  strDir = Utils.GetPath(dictTemplateInfo("Final"))

  If IsItThere(strDir) = False Then
    MkDir (strDir)
    IsTemplateThere = False
    logString = Now & " -- Creating template directory."
  Else
  ' Check if template file exists
    If IsItThere(dictTemplateInfo("Final")) = False Then
      IsTemplateThere = False
      logString = Now & " -- " & FileName & " doesn't exist in " & strDir
    Else
      IsTemplateThere = True
      logString = Now & " -- " & FileName & " already exists."
    End If
  End If

  LogInformation dictTemplateInfo("Log"), logString
End Function

Private Function NeedUpdate(FileName As String) As Boolean
' FileName should be to TEMPLATE (.dotm) file
  Dim dictTemplateFile As Dictionary
  Set dictTemplateFile = FileInfo(FileName)

'------------------------- Get installed version number -----------------------------------
  Dim logString As String

' Get version number of installed template
  Dim strInstalledVersion As String

  If IsItThere(dictTemplateFile("Final")) = True Then

    #If Mac Then
      Call OpenDocMac(dictTemplateFile("Final"))
    #Else
      Call OpenDocPC(dictTemplateFile("Final"))
    #End If

    strInstalledVersion = Documents(dictTemplateFile("Final")).CustomDocumentProperties("Version")
    Documents(dictTemplateFile("Final")).Close SaveChanges:=wdDoNotSaveChanges
    logString = Now & " -- installed version is " & strInstalledVersion
  Else
    strInstalledVersion = "0"     ' Template is not installed
    logString = Now & " -- No template installed, version number is 0."
  End If

  LogInformation dictTemplateFile("Log"), logString

'------------------------- Try to get current version's number from Confluence
  Dim strVersion As String
  Dim dictVersionFile As Dictionary

  Dim strMacDocs As String

  strVersion = Utils.GetFileNameOnly(FileName) & ".txt"
  Set dictVersionFile = FileInfo(strVersion)

' If False, error in download; user was notified in DownloadFromGithub function
  If DownloadFromGithub(FileName:=strVersion) = False Then
    NeedUpdate = False
    Exit Function
  End If

'-------------------- Get version number of current template ---------------------
  If IsItThere(dictVersionFile("Final")) = True Then
    NeedUpdate = True
    Dim strCurrentVersion As String
    strCurrentVersion = ReadTextFile(Path:=dictVersionFile("Final"), FirstLineOnly:=True)

  ' git converts all line endings to LF which messes up PC, and I don't want to deal
  ' with it so we'll just remove everything
    If InStr(strCurrentVersion, vbCrLf) > 0 Then
      strCurrentVersion = Replace(strCurrentVersion, vbCrLf, "")
    ElseIf InStr(strCurrentVersion, vbLf) > 0 Then
      strCurrentVersion = Replace(strCurrentVersion, vbLf, "")
    ElseIf InStr(strCurrentVersion, vbCr) > 0 Then
      strCurrentVersion = Replace(strCurrentVersion, vbCr, "")
    End If

    logString = Now & " -- Current version is " & strCurrentVersion
  Else
    NeedUpdate = False
    logString = Now & " -- Download of version file for " & FileName & " failed."
  End If

  LogInformation dictVersionFile("Log"), logString

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

  LogInformation dictVersionFile("Log"), logString

End Function

Private Sub OpenDocMac(FilePath As String)
        Documents.Open FileName:=FilePath, ReadOnly:=True ', Visible:=False      'Mac Word 2011 doesn't allow Visible as an argument :(
End Sub

Private Sub OpenDocPC(FilePath As String)
        Documents.Open FileName:=FilePath, ReadOnly:=True, Visible:=False      'Win Word DOES allow Visible as an argument :)
End Sub

Private Function FullURL(FileName As String) As String
' Takes a file name as an argument and returns the URL to that file ON GITHUB
' TODO: Create from config file
  Dim strBaseUrl As String
  Dim strBranch As String
  Dim strNameOnly As String
  Dim strRepo As String
  Dim strSubfolder As String
  Dim strFilePath As String

  strBaseUrl = "https://raw.githubusercontent.com/macmillanpublishers"
' Strip extension, bc some files have related version file w/ same name, diff ext
  strNameOnly = Utils.GetFileNameOnly(FileName)

  Select Case strNameOnly
    Case "macmillan"
      strRepo = "Word-template_assets"
      strSubfolder = "StyleTemplate_auto-generate"
    Case "macmillan_NoColor"
      strRepo = "Word-template_assets"
      strSubfolder = "StyleTemplate_auto-generate"
    Case "macmillan_CoverCopy"
      strRepo = "Word-template_assets"
      strSubfolder = vbNullString
    Case "Styles_Bookmaker"
      strRepo = "Word-template_assets"
      strSubfolder = vbNullString
    Case "Word-template"
      strRepo = "Word-template"
      strSubfolder = vbNullString
    Case "GtUpdater"
      strRepo = "Word-template"
      strSubfolder = vbNullString
    Case "section_start_rules"
      strRepo = "bookmaker_validator"
      strSubfolder = vbNullString
    Case "vba_style_config"
      strRepo = "Word-template_assets"
      strSubfolder = "StyleTemplate_auto-generate"
  End Select

' Get branch based on repo for now, so don't all have to be on same branch
' TODO: Read branch for each file from a config
  strBranch = WT_Settings.DownloadBranch(Repo:=strRepo)
  
' Test if strSubfolder exists, and if so combine to create full file path.
' If we combine below and an item is blank, will get a double separator
  If strSubfolder <> vbNullString Then
    strFilePath = strSubfolder & "/"
  End If
  
  strFilePath = strFilePath & FileName
  
  ' put it all together
  FullURL = strBaseUrl & "/" & strRepo & "/" & strBranch & "/" & strFilePath
End Function


Private Function ImportVariable(strFile As String) As String
 
    Open strFile For Input As #1
    Line Input #1, ImportVariable
    Close #1
 
End Function


Private Sub CloseOpenDocs()

  '-------------Check for/close open documents---------------------------------
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
          'DebugPrint doc.Name
          If doc.Name <> strInstallerName Then       'But don't close THIS document
            doc.Save   'separate step to trigger Save As prompt for previously unsaved docs
            doc.Close
          End If
        On Error GoTo 0
      Next doc
    End If
  End If
    
End Sub

' ===== TmpDir ================================================================
' Returns path to TEMP directory, no trailing path separator

Private Function TmpDir() As String
  Dim strTmpDir As String
  #If Mac Then
    strTmpDir = MacScript("path to temporary items as string")
  #Else
    strTmpDir = Environ("TEMP")
  #End If
' Remove trailing path separator, if any
  If Right(strTmpDir, 1) = Application.PathSeparator Then
    strTmpDir = Left(strTmpDir, Len(strTmpDir) - 1)
  End If
  TmpDir = strTmpDir
End Function

' ===== FileInfo ==============================================================
' Returns dictionary with paths to final file location and other things for
' downloading.

' PARAMS
' FileName: File name with extension that you want to download.

Public Function FileInfo(FileName As String) As Dictionary
  Dim dictFileInfo As Dictionary
  Set dictFileInfo = New Dictionary
  
  Dim strStyleDir As String
  Dim strTmpDir As String
  Dim strBaseName As String
  Dim strSep As String

' Becasue I am too lazy to write this whole thing later
  strSep = Application.PathSeparator
  
' GtUpdater.dotm only file to go in Startup
  If FileName = "GtUpdater.dotm" Then
    strStyleDir = Application.StartupPath
  Else
    strStyleDir = StyleDir()
  End If

' Create directory if it doesn't exist already
  If Utils.IsItThere(strStyleDir) = False Then
    MkDir strStyleDir
  End If

  strTmpDir = TmpDir()
  strBaseName = Utils.GetFileNameOnly(FileName)
  
  With dictFileInfo
    .Add "Final", strStyleDir & strSep & FileName
    .Add "Tmp", strTmpDir & strSep & FileName
    .Add "Log", strStyleDir & strSep & "log" & strSep & strBaseName & "_updates.log"
  End With
  
  Set FileInfo = dictFileInfo

End Function
