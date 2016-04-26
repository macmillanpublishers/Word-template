VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MacFile_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' Class for managing the files that constitute the Macmillan Tools suite
' By Erica Warren - erica.warren@macmillan.com

' DEPENDENCIES:
' * config.json file with specific structure, containing info about files,
'   download source, local directory structure
' * Paths_.cls static class module
' * https://github.com/VBA-tools/VBA-Dictionary
' * https://github.com/VBA-tools/VBA-JSON
' =============================================================================


'' *****************************************************************************
''           CLASS ENUMERATIONS
'' *****************************************************************************

'' ===== MacFileGroup ==========================================================
'' This enum should include the first-level keys in the "files" object in
'' the config JSON file. Must be set to powers of two so we can use bitwise
'' flags to combine groups.
'
'Public Enum MacFileGroup
'    configFiles = 2 ^ 0     ' 1
'    startupFiles = 2 ^ 1    ' 2
'    macroFiles = 2 ^ 2      ' 4
'    styleFiles = 2 ^ 3      ' 8
'    assetFiles = 2 ^ 4      ' 16
'    updateCheck = macroFiles + styleFiles
'    wordDocs = startupFiles + macroFiles + styleFiles
'    installFiles = startupFiles + macroFiles + styleFiles + assetFiles
'    allFiles = configFiles + noConfig
'End Enum


' *****************************************************************************
'           CLASS VARIABLES
' *****************************************************************************

Private p_GroupName As String
Private p_SpecificFile As String
Private p_GroupJSON As Dictionary
Private p_FileJSON As Dictionary
Private p_ConfigVersion As Dictionary
Private p_LocalVersion As Dictionary


' *****************************************************************************
'           PUBLIC PROPERTIES
' *****************************************************************************

' ===== GroupName =============================================================
' Group here is in reference to the first level of keys in the "file" object in
' JSON file, grouping files with similar purposes together.

Public Property Let GroupName(GroupName_value As String)
    If Paths_.FullJSON("files").Exists(GroupName_value) Then
        p_GroupName = GroupName_value
    Else
        MsgBox "That's not a valid option for the GroupName property."
    End If
End Property

Public Property Get GroupName() As String
    If p_GroupName <> vbNullString Then
        GroupName = p_GroupName
    Else
        MsgBox "You must set the GroupName property first!"
    End If
End Property

' ===== SpecificFile ==========================================================
' Specific meaning "category" or "role"; one level below GroupName in the JSON
' file; i.e., what we are calling the file based on what it does in the project
' (not its Name, which can change). Options for specific file are limited to
' the keys of the various objects in the GroupNames, and each must me unique.
' Must also Set the FileJSON property now so we can access it later.
' The SpecificFile property needs to be set before you can access other info.

Public Property Let SpecificFile(SpecificFile_value As String)
    If GroupJSON.Exists(SpecificFile_value) Then
        p_SpecificFile = SpecificFile_value
    Else
        MsgBox "That's not a valid option for this property."
    End If
End Property

Public Property Get SpecificFile() As String
    If p_SpecificFile <> vbNullString Then
        SpecificFile = p_SpecificFile
    Else
        MsgBox "You must set the SpecificFile property first!"
    End If
End Property

' ===== FullName ==============================================================
' Full file name including extension, but not any path info

Public Property Get FullName() As String
    FullName = Name & "." & Extension
End Property

' ===== FullLocalPath =========================================================
' Returns string of the full path to the file, including root and full file
' name with extension.

Public Property Get FullLocalPath() As String
    FullLocalPath = Paths_.LocalPath(GroupName) & _
        Application.PathSeparator & FullName
End Property

' ===== Exists ================================================================
' Tests whether the file exists at its FullLocalPath

Public Property Get Exists() As Boolean
    Exists = SharedMacros.IsItThere(Path:=FullLocalPath)
End Property

' ===== Running ===============================================================
' Tests whether the file in question is the one that's currently running this
' code, so we don't, like, delete ourselves or something.

Public Property Get Running() As Boolean
    If ThisDocument.Name = FullName Then
        Running = True
    Else
        Running = False
    End If
End Property

' ===== UpToDate ==============================================================
' Checks if the local file version number is up to date with the config file
' version number. Only returns False if the local version is LOWER; if the
' local version is HIGHER than the config, returns True. Compares each part of
' the version number (Major.Minor.Patch) in turn.

Public Property Get UpToDate() As Boolean
    If ConfigVersion("major") > LocalVersion("major") Then
        UpToDate = False
    ElseIf ConfigVersion("minor") > LocalVersion("minor") Then
        UpToDate = False
    ElseIf ConfigVersion("patch") > LocalVersion("patch") Then
        UpToDate = False
    Else
        UpToDate = True
    End If
End Property


' *****************************************************************************
'           PRIVATE PROPERTIES
' *****************************************************************************

' ===== GroupJSON =============================================================
' Dictionary object of the GroupName level in the JSON file. Not entirely sure
' we definitely need this here, but then I referenced it in two places so I
' figured why not.

Private Property Get GroupJSON() As Dictionary
    If p_GroupJSON Is Nothing Then
        Set p_GroupJSON = Paths_.FullJSON("files")(GroupName)
    End If
    Set GroupJSON = p_GroupJSON
End Property

' ===== FileJSON ==============================================================
' Dictionary object of the info for the specific file.

Private Property Get FileJSON() As Dictionary
    If p_FileJSON Is Nothing Then
        Set p_FileJSON = GroupJSON(SpecificFile)
    End If
    Set FileJSON = p_FileJSON
End Property

' ===== Name ==================================================================
' Name of file as a string, WITHOUT extension. Keep name and extension separate
' because some things (like log file names) require just the name w/o extension.

Private Property Get Name() As String
    Name = FileJSON("fileName")
End Property

' ===== Extension =============================================================
' File extension as a string, NOT including period.

Private Property Get Extension() As String
    Extension = FileJSON("extension")
End Property

' ===== ConfigVersion ==========================================================
' Version of the file as noted in the config.JSON file. Note that Patch can be
' null but Major and Minor cannot.

Private Property Get ConfigVersion() As Dictionary
    If p_ConfigVersion Is Nothing Then
        Set p_ConfigVersion = FileJSON("version")
    End If
    Set ConfigVersion = p_ConfigVersion
End Property

' ===== LocalVersion ==========================================================
' For Word documents/templates: version of the file as noted in the Custom
' Document Properties. Note, stored as three separate properties for Major,
' Minor, and Patch. Patch may be null. Document must be open (though might be a
' way to manage this by unzipping the file and reading XML directly?).

Private Property Get LocalVersion() As Dictionary
    ' Word files: .docx, .dotx, .docm, .dotm
    If Left(Extension, 2) = "do" And _
        InStr("ct", Mid(Extension, 3, 1)) > 0 And _
        InStr("mx", Right(Extension, 1)) > 0 Then
        
        If p_LocalVersion Is Nothing Then
            Set p_LocalVersion = New Dictionary
            Dim objDoc As Document
            Set objDoc = Documents.Open _
                (FileName:=FullLocalPath, ReadOnly:=True, Visible:=False)
            p_LocalVersion.Item("major") = objDoc.CustomDocumentProperties("Major")
            p_LocalVersion.Item("minor") = objDoc.CustomDocumentProperties("Minor")
            p_LocalVersion.Item("patch") = objDoc.CustomDocumentProperties("Patch")
        Else
            Set LocalVersion = p_LocalVersion
        End If
    Else
        Set LocalVersion = ConfigVersion
    End If
End Property

' ===== FullLogPath ===========================================================
' Full path to the log file, including file name and extension.

Private Property Get FullLogPath() As String
    FullLogPath = Paths_.LocalPath(GroupName:="logFiles") & _
        Application.PathSeparator & Name & "_updates.log"
End Property

' ===== FullTempPath ==========================================================
' Full path to the file in the temp dir. NOT verifying if it's there, because
' we need to reference this before we can create it with ParseDoc, maybe.

Private Property Get FullTempPath() As String
    FullTempPath = Paths_.TempPath & Application.PathSeparator & FullName
End Property


' *****************************************************************************
'           PUBLIC METHODS
' *****************************************************************************

' ===== AssignFile ============================================================
' To access all of the properties of the object, both GroupName & SpecificFile
' properties must be set. To make this easier, this method does both at once.

Public Sub AssignFile(strGroupName As String, strSpecificFileType As String)
    GroupName = strGroupName
    SpecificFile = strSpecificFileType
End Sub

' ===== DeleteFile ============================================================
' Deletes the file in question, as long as it's not currently running this code

Public Sub DeleteFile()
    If Running = False Then
        If Exists = True Then
            Kill FullLocalPath
        End If
    Else
        MsgBox "You can't delete a file that is running a macro."
    End If
End Sub

' ===== NotifyComplete ========================================================
' Notifies the user that the download is complete. Outside Download method cuz
' we only want to notify once even if we're downloading multiple files; also
' some downloads are silent (e.g., Castoff values), though we should always
' notify when we're downloading files that may contain updated functionality.
' Also NOT in MacFileBunch_ because if the main macro template itself needs to
' be updated, it will launch the updater template, which will only download a
' single file but which will notify for all downloads.

Public Sub NotifyComplete()
    MsgBox "All downloads are complete. Yay!"
End Sub

'' ===== Download ==============================================================
'' Downloads the file, which must be posted at the URL returned by the
'' DownloadURL property. Returns False if there was an error, and it notifies
'' the user of the error in the function -- but should probably move this to a
'' global error handler class at some point.
'
'Public Function Download() As Boolean
''
''    Dim logString As String
''    Dim strTmpPath As String
''    Dim strBashTmp As String
''    Dim strErrMsg As String
''
''    strTmpPath = Paths_.TempPath & Application.PathSeparator & FullName
''
''    ' Mac can't use "MSXML2.XMLHTTP.3.0" lib, so we'll use curl
''    #If Mac Then
''        strBashTmp = Replace(Right(strTmpPath, Len(strTmpPath) - (InStr _
''            (strTmpPath, ":") - 1)), ":", "/")
''
''        'check for network
''        If ShellAndWaitMac("ping -o google.com &> /dev/null ; echo $?") _
''            <> 0 Then   'can't connect to internet
''            logString = "Tried update; unable to connect to network."
''            WriteToLog LogMessage:=logString
''            strErrMsg = "There was an error trying to download the " & _
''                "Macmillan template." & vbNewLine & vbNewLine & "Please " & _
''                "check your internet connection or contact " & _
''                Organization_.HelpEmail & " for help."
''            MsgBox strErrMsg, vbCritical, "Error 1: Connection error (" & _
''                FullName & ")"
''            Download = False
''            Exit Function
''        Else 'internet is working, download file
''            'Make sure file is there
''            Dim httpStatus As Long
''            httpStatus = ShellAndWaitMac _
''                ("curl -s -o /dev/null -w '%{http_code}' " & DownloadURL)
''
''            If httpStatus = 200 Then                    ' File is there
''                'Now delete file if already there, then download new file
''                ShellAndWaitMac ("rm -f " & strBashTmp & " ; curl -o " & _
''                    strBashTmp & " " & DownloadURL)
''            ElseIf httpStatus = 404 Then            ' 404 = page not found
''                logString = "404 File not found. Cannot download file."
''                WriteToLog LogMessage:=logString
''                strErrMsg = "It looks like that file isn't available for " & _
''                    "download." & vbNewLine & vbNewLine & "Please contact " & _
''                    Organization_.HelpEmail & " for help."
''                MsgBox strErrMsg, vbCritical, "Error 7: File not found (" & _
''                    FullName & ")"
''                Download = False
''                Exit Function
''            Else
''                logString = "Http status is " & httpStatus & ". Cannot " & _
''                    "download file."
''                WriteToLog LogMessage:=logString
''                strErrMsg = "There was an error trying to download the " & _
''                    "Macmillan templates." & vbNewLine & vbNewLine & "Please" _
''                    & " check your internet connection or contact " & _
''                    Organization_.HelpEmail & " for help."
''                MsgBox strErrMsg, vbCritical, "Error 2: Http status " & _
''                    httpStatus & " (" & FullName & ")"
''                Download = False
''                Exit Function
''            End If
''
''        End If
''    #Else   ' PC download
''
''        'Check if file is already in tmp dir, delete if yes
''        If IsItThere(strTmpPath) = True Then
''            Kill strTmpPath
''        End If
''
''        'try to download the file from Public Confluence page
''        Dim WinHttpReq As Object
''        Dim oStream As Object
''
''        'Attempt to download file
''        On Error Resume Next
''            Set WinHttpReq = CreateObject("MSXML2.XMLHTTP.3.0")
''            WinHttpReq.Open "GET", DownloadURL, False
''            WinHttpReq.Send
''
''                ' Exit sub if error in connecting to website
''                If Err.Number <> 0 Then 'HTTP request is not OK
''                    'Debug.Print WinHttpReq.Status
''                    logString = Now & " -- could not connect to Confluence site: Error " & Err.Number
''                    WriteToLog LogMessage:=logString
''                    strErrMsg = "There was an error trying to download the Macmillan template." & vbNewLine & vbNewLine & _
''                        "Please check your internet connection or contact workflows@macmillan.com for help."
''                    MsgBox strErrMsg, vbCritical, "Error 1: Connection error (" & FullName & ")"
''                    Download = False
''                    On Error GoTo 0
''                    Exit Function
''                End If
''        On Error GoTo 0
''
''        'Debug.Print "Http status for " & FullName & ": " & WinHttpReq.Status
''        If WinHttpReq.Status = 200 Then  ' 200 = HTTP request is OK
''
''            'if connection OK, download file to temp dir
''            DownloadURL = WinHttpReq.ResponseBody
''            Set oStream = CreateObject("ADODB.Stream")
''            oStream.Open
''            oStream.Type = 1
''            oStream.Write WinHttpReq.ResponseBody
''            oStream.SaveToFile strTmpPath, 2 ' 1 = no overwrite, 2 = overwrite
''            oStream.Close
''            Set oStream = Nothing
''            Set WinHttpReq = Nothing
''        ElseIf WinHttpReq.Status = 404 Then ' 404 = file not found
''            logString = Now & " -- 404 File not found. Cannot download file."
''            WriteToLog LogMessage:=logString
''            strErrMsg = "It looks like that file isn't available for download." & vbNewLine & vbNewLine & _
''                "Please contact workflows@macmillan.com for help."
''            MsgBox strErrMsg, vbCritical, "Error 7: File not found (" & FullName & ")"
''            Download = False
''            Exit Function
''        Else
''            logString = Now & " -- Http status is " & WinHttpReq.Status & ". Cannot download file."
''            WriteToLog LogMessage:=logString
''            strErrMsg = "There was an error trying to download the Macmillan templates." & vbNewLine & vbNewLine & _
''                "Please check your internet connection or contact workflows@macmillan.com for help."
''            MsgBox strErrMsg, vbCritical, "Error 2: Http status " & WinHttpReq.Status & " (" & FullName & ")"
''            Download = False
''            Exit Function
''        End If
''    #End If
''
''    'Error if download was not successful
''    If IsItThere(strTmpPath) = False Then
''        logString = Now & " -- " & FullName & " file download to Temp was not successful."
''        WriteToLog LogMessage:=logString
''        strErrMsg = "There was an error downloading the Macmillan template." & vbNewLine & _
''            "Please contact workflows@macmillan.com for assitance."
''        MsgBox strErrMsg, vbCritical, "Error 3: Download failed (" & FullName & ")"
''        Download = False
''        On Error GoTo 0
''        Exit Function
''    Else
''        logString = Now & " -- " & FullName & " file download to Temp was successful."
''        WriteToLog LogMessage:=logString
''    End If
''
''
''
''    'If file exists already, log it and delete it
''    If IsItThere(strFinalPath) = True Then
''
''        logString = Now & " -- Previous version file in final directory."
''        WriteToLog LogMessage:=logString
''
''        ' get file extension
''        Dim strExt As String
''        strExt = Right(strFinalPath, InStrRev(StrReverse(strFinalPath), "."))
''
''        ' can't delete template if it's installed as an add-in
''        If InStr(strExt, "dot") > 0 Then
''            On Error Resume Next        'Error = add-in not available, don't need to uninstall
''                AddIns(strFinalPath).Installed = False
''            On Error GoTo 0
''        End If
''
''        ' Test if dir is read only
''        If IsReadOnly(FinalDir) = True Then ' Dir is read only
''            logString = Now & " -- old " & FullName & " file is read only, can't delete/replace. " _
''                & "Alerting user."
''            WriteToLog LogMessage:=logString
''            strErrMsg = "The installer doesn't have permission. Please conatct workflows" & _
''                "@macmillan.com for help."
''            MsgBox strErrMsg, vbCritical, "Error 8: Permission denied (" & FullName & ")"
''            Download = False
''            On Error GoTo 0
''            Exit Function
''        Else
''            On Error Resume Next
''                Kill strFinalPath
''
''                If Err.Number = 70 Then         'File is open and can't be replaced
''                    logString = Now & " -- old " & FullName & " file is open, can't delete/replace. Alerting user."
''                    WriteToLog LogMessage:=logString
''                    strErrMsg = "Please close all other Word documents and try again."
''                    MsgBox strErrMsg, vbCritical, "Error 4: Previous version removal failed (" & FullName & ")"
''                    Download = False
''                    On Error GoTo 0
''                    Exit Function
''                End If
''            On Error GoTo 0
''        End If
''    Else
''        logString = Now & " -- No previous version file in final directory."
''        WriteToLog LogMessage:=logString
''    End If
''
''    'If delete was successful, move downloaded file to final directory
''    If IsItThere(strFinalPath) = False Then
''        logString = Now & " -- Final directory clear of " & FullName & " file."
''        WriteToLog LogMessage:=logString
''
''        ' move template to final directory
''        Name strTmpPath As strFinalPath
''
''        'Mac won't load macros from a template downloaded from the internet to Startup.
''        'Need to send these commands for it to work, see Confluence
''        ' Do NOT use open/save as option, this removes customUI which creates Mac Tools toolbar later
''        #If Mac Then
''            If InStr(1, FullName, ".dotm") Then
''            Dim strCommand As String
''            strCommand = "do shell script " & Chr(34) & "xattr -wx com.apple.FinderInfo \" & Chr(34) & _
''                "57 58 54 4D 4D 53 57 44 00 10 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00\" & _
''                Chr(34) & Chr(32) & Chr(34) & " & quoted form of POSIX path of " & Chr(34) & strFinalPath & Chr(34)
''                'Debug.Print strCommand
''                MacScript (strCommand)
''            End If
''        #End If
''
''    Else
''        logString = Now & " -- old " & FullName & " file not cleared from Final directory."
''        WriteToLog LogMessage:=logString
''        strErrMsg = "There was an error installing the Macmillan template." & vbNewLine & _
''            "Please close all other Word documents and try again, or contact workflows@macmillan.com."
''        MsgBox strErrMsg, vbCritical, "Error 5: Previous version uninstall failed (" & FullName & ")"
''        Download = False
''        On Error GoTo 0
''        Exit Function
''    End If
''
''    'If move was successful, yay! Else, :(
''    If IsItThere(strFinalPath) = True Then
''        logString = Now & " -- " & FullName & " file successfully saved to final directory."
''        WriteToLog LogMessage:=logString
''    Else
''        logString = Now & " -- " & FullName & " file not saved to final directory."
''        WriteToLog LogMessage:=logString
''        strErrMsg = "There was an error installing the Macmillan template." & vbNewLine & vbNewLine & _
''            "Please cotact workflows@macmillan.com for assistance."
''        MsgBox strErrMsg, vbCritical, "Error 6: Installation failed (" & FullName & ")"
''        Download = False
''        On Error GoTo 0
''        Exit Function
''    End If
''
''    'Cleanup: Get rid of temp file if downloaded correctly
''    If IsItThere(strTmpPath) = True Then
''        Kill strTmpPath
''    End If
''
''    ' Disable Startup add-ins so they don't launch right away and mess of the code that's running
''    If InStr(1, LCase(strFinalPath), LCase("startup"), vbTextCompare) > 0 Then         'LCase because "startup" was staying in all caps for some reason, UCase wasn't working
''        On Error Resume Next                                        'Error = add-in not available, don't need to uninstall
''            AddIns(strFinalPath).Installed = False
''        On Error GoTo 0
''    End If
''
''    Download = True
'
'End Function
'

' *****************************************************************************
'           PRIVATE METHODS
' *****************************************************************************

' ===== StringVersion =========================================================
' returns full version number as a string. May need to make Public? Takes an
' argument so we can use it to return either the Config or Local version. Note
' "patch" may be Null

Private Function StringVersion(dictVersion As Dictionary) As String
    Dim strVersionReturn As String
    strVersionReturn = CStr(dictVersion("major")) & "." & CStr(dictVersion("minor"))

    If CStr(dictVersion("patch")) <> vbNullString Then
        strVersionReturn = strVersionReturn & "." & CStr(dictVersion("patch"))
    End If

    StringVersion = strVersionReturn
End Function

' ===== WriteToLog ============================================================
' Writes line to log for the file. LogMessage only needs text, timestamp will
' be added in this method. May make public later if we end up using this more
' more than just download logs.

Private Sub WriteToLog(LogMessage As String)
    Dim strLogMessage As String
    Dim FileNum As Integer

    strLogMessage = Now & " -- " & LogMessage
    FileNum = FreeFile ' next file number

    Open LogFile For Append As #FileNum ' creates the file if it doesn't exist
    Print #FileNum, strLogMessage ' write information at the end of the text file
    Close #FileNum ' close the file
End Sub


' *****************************************************************************
'       CLASS EVENTS
' *****************************************************************************

' Object properties will persist after the calling procedure is ended, even if
' you set the MacFile_ object to Nothing at the end of that sub.

Sub Class_Terminate()
    If Not p_GroupJSON Is Nothing Then
        Set p_GroupJSON = Nothing
    End If
    
    If Not p_FileJSON Is Nothing Then
        Set p_FileJSON = Nothing
    End If
    
    If Not p_ConfigVersion Is Nothing Then
        Set p_ConfigVersion = Nothing
    End If
    
    If Not p_LocalVersion Is Nothing Then
        Set p_LocalVersion = Nothing
    End If
End Sub
