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


' *****************************************************************************
'           CLASS VARIABLES
' *****************************************************************************

Private p_GroupName As String
Private p_SpecificFile As String
Private p_GroupJSON As Dictionary
Private p_FileJSON As Dictionary
Private p_ConfigVersion As Dictionary
Private p_LocalVersion As Dictionary
Private p_Exists As Boolean
Private p_Extension As String
Private p_FullLocalPath As String
Private p_FullLogPath As String
Private p_FullName As String
Private p_FullTempPath As String
Private p_Name As String

Private Const strModuleName As String = "MacFile_."

' *****************************************************************************
'           PUBLIC PROPERTIES
' *****************************************************************************

' ===== GroupName =============================================================
' Group here is in reference to the first level of keys in the "file" object in
' JSON file, grouping files with similar purposes together.

Public Property Let GroupName(GroupName_value As String)
    On Error GoTo GroupNameLetError
    If Paths_.FullJSON("files").Exists(GroupName_value) Then
        p_GroupName = GroupName_value
    Else
        Err.Raise MacError.err_GroupNameInvalid
    End If
GroupNameLetFinish:
    Exit Property

GroupNameLetError:
    Err.Source = Err.Source & strModuleName & "GroupName.Let"
    If SharedMacros_.ErrorChecker(Err, GroupName_value) = False Then
        Resume
    Else
        Resume GroupNameLetFinish
    End If
End Property

Public Property Get GroupName() As String
    On Error GoTo GroupNameGetError
    If p_GroupName <> vbNullString Then
        GroupName = p_GroupName
    Else
        Err.Raise MacError.err_GroupNameNotSet
    End If
GroupNameGetFinish:
    Exit Property

GroupNameGetError:
    Err.Source = Err.Source & strModuleName & "GroupName.Get"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume GroupNameGetFinish
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
    On Error GoTo SpecificFileLetError
    If GroupJSON.Exists(SpecificFile_value) Then
        p_SpecificFile = SpecificFile_value
    Else
        Err.Raise MacError.err_SpecificFileInvalid
    End If
SpecificFileLetFinish:
    Exit Property

SpecificFileLetError:
    Err.Source = Err.Source & strModuleName & "SpecificFile.Let"
    If SharedMacros_.ErrorChecker(Err, SpecificFile_value) = False Then
        Resume
    Else
        Resume SpecificFileLetFinish
    End If
End Property

Public Property Get SpecificFile() As String
    On Error GoTo SpecifiFileGetError
    If p_SpecificFile <> vbNullString Then
        SpecificFile = p_SpecificFile
    Else
        Err.Raise MacError.err_SpecificFileNotSet
    End If
SpecificFileGetFinish:
    Exit Property

SpecificFileGetError:
    Err.Source = Err.Source & strModuleName & "SpecificFile.Get"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume SpecificFileGetFinish
    End If
End Property

' ===== FullName ==============================================================
' Full file name including extension, but not any path info. Read only.

Public Property Get FullName() As String
    On Error GoTo FullNameGetError
    If p_FullName = vbNullString Then
        p_FullName = Name & "." & Extension
    End If
    FullName = p_FullName
FullNameGetFinish:
    Exit Property
    
FullNameGetError:
    Err.Source = Err.Source & strModuleName & "FullName.Get"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume FullNameGetFinish
    End If
End Property

' ===== FullLocalPath =========================================================
' Returns string of the full path to the file, including root and full file
' name with extension. Read only.

Public Property Get FullLocalPath() As String
    On Error GoTo FullLocalPathGetError
    If p_FullLocalPath = vbNullString Then
        p_FullLocalPath = Paths_.LocalPath(GroupName) & _
            Application.PathSeparator & FullName
        ' Check if final dir is read-only
        If SharedMacros_.IsReadOnly(p_FullLocalPath) = True Then
            Err.Raise MacError.err_LocalReadOnly
        End If
    End If
    FullLocalPath = p_FullLocalPath
FullLocalPathGetFinish:
    Exit Property
    
FullLocalPathGetError:
    Err.Source = Err.Source & strModuleName & "FullLocalPath.Get"
    If SharedMacros_.ErrorChecker(Err, p_FullLocalPath) = False Then
        Resume
    Else
        Resume FullLocalPathGetFinish
    End If
End Property

' ===== Exists ================================================================
' Tests whether the file exists at its FullLocalPath. No private variable; we
' always want to check in case it's deleted at runtime. Read only.

Public Property Get Exists() As Boolean
    On Error GoTo ExistsGetError
    Exists = SharedMacros_.IsItThere(Path:=FullLocalPath)
ExistsGetFinish:
    Exit Property
    
ExistsGetError:
    Err.Source = Err.Source & strModuleName & "Exists.Get"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume ExistsGetFinish
    End If
End Property

' ===== Running ===============================================================
' Tests whether the file in question is the one that's currently running this
' code, so we don't, like, delete ourselves or something. Always check. Read
' only.

Public Property Get Running() As Boolean
    On Error GoTo RunningGetError
    Dim strThisName As String
    strThisName = ThisDocument.Name
    If strThisName = FullName Then
        Running = True
    Else
        Running = False
    End If
RunningGetFinish:
    Exit Property

RunningGetError:
    Err.Source = Err.Source & strModuleName & "Running.Get"
    If SharedMacros_.ErrorChecker(Err, strThisName) = False Then
        Resume
    Else
        Resume RunningGetFinish
    End If
End Property

' ===== UpToDate ==============================================================
' Checks if the local file version number is up to date with the config file
' version number. Only returns False if the local version is LOWER; if the
' local version is HIGHER than the config, returns True. Compares each part of
' the version number (Major.Minor.Patch) in turn. Read only.

Public Property Get UpToDate() As Boolean
    On Error GoTo UpToDateGetError
    If ConfigVersion("major") > LocalVersion("major") Then
        UpToDate = False
    ElseIf ConfigVersion("minor") > LocalVersion("minor") Then
        UpToDate = False
    ElseIf ConfigVersion("patch") > LocalVersion("patch") Then
        UpToDate = False
    Else
        UpToDate = True
    End If
UpToDateGetFinish:
    Exit Property
    
UpToDateGetError:
    Err.Source = Err.Source & strModuleName & "UpToDate.Get"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume UpToDateGetFinish
    End If
End Property


' *****************************************************************************
'           PRIVATE PROPERTIES
' *****************************************************************************

' ===== GroupJSON =============================================================
' Dictionary object of the GroupName level in the JSON file. Not entirely sure
' we definitely need this here, but then I referenced it in two places so I
' figured why not. Read only.

Private Property Get GroupJSON() As Dictionary
    On Error GoTo GroupJSONGetError
    If p_GroupJSON Is Nothing Then
        Set p_GroupJSON = Paths_.FullJSON("files")(GroupName)
    End If
    Set GroupJSON = p_GroupJSON
GroupJSONGetFinish:
    Exit Property

GroupJSONGetError:
    Err.Source = Err.Source & strModuleName & "GroupJSON.Get"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume GroupJSONGetFinish
    End If
End Property

' ===== FileJSON ==============================================================
' Dictionary object of the info for the specific file. Read only.

Private Property Get FileJSON() As Dictionary
    On Error GoTo FileJSONGetError
    If p_FileJSON Is Nothing Then
        Set p_FileJSON = GroupJSON(SpecificFile)
    End If
    Set FileJSON = p_FileJSON
FileJSONGetFinish:
    Exit Property
    
FileJSONGetError:
    Err.Source = Err.Source & strModuleName & "FileJSON.Get"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume FileJSONGetFinish
    End If
End Property

' ===== Name ==================================================================
' Name of file as a string, WITHOUT extension. Keep name and extension separate
' because some things (like log file names) require just the name w/o extension.
' Read only.

Private Property Get Name() As String
    On Error GoTo NameGetError
    If p_Name = vbNullString Then
        p_Name = FileJSON("fileName")
    End If
    Name = p_Name
NameGetFinish:
    Exit Property

NameGetError:
    Err.Source = Err.Source & strModuleName & "Name.Get"
    If SharedMacros_.ErrorChecker(Err, p_Name) = False Then
        Resume
    Else
        Resume NameGetFinish
    End If
End Property

' ===== Extension =============================================================
' File extension as a string, NOT including period. Read only.

Private Property Get Extension() As String
    On Error GoTo ExtensionGetError
    If p_Extension = vbNullString Then
        p_Extension = FileJSON("extension")
    End If
    Extension = p_Extension
ExtensionGetFinish:
    Exit Property
    
ExtensionGetError:
    Err.Source = Err.Source & strModuleName & "Extension.Get"
    If SharedMacros_.ErrorChecker(Err, p_Extension) = False Then
        Resume
    Else
        Resume ExtensionGetFinish
    End If
End Property

' ===== ConfigVersion ==========================================================
' Version of the file as noted in the config.JSON file. Note that Patch can be
' null but Major and Minor cannot. Read only.

Private Property Get ConfigVersion() As Dictionary
    On Error GoTo ConfigVersionGetError
    If p_ConfigVersion Is Nothing Then
        Set p_ConfigVersion = FileJSON("version")
    End If
    Set ConfigVersion = p_ConfigVersion
ConfigVersionGetFinish:
    Exit Property
    
ConfigVersionGetError:
    Err.Source = Err.Source & strModuleName & "ConfigVersion.Get"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume ConfigVersionGetFinish
    End If
End Property

' ===== LocalVersion ==========================================================
' For Word documents/templates: version of the file as noted in the Custom
' Document Properties. Stored as three separate properties for Major, Minor,
' and Patch. Patch may be null. Document must be open (though might be a way to
' manage this by unzipping the file and reading XML directly?). Read only.

' Will need to create a Property Let for Utilities updating stuff.

Private Property Get LocalVersion() As Dictionary
    On Error GoTo LocalVersionGetError
    If p_LocalVersion Is Nothing Then
        ' Word files: .docx, .dotx, .docm, .dotm
        If Left(Extension, 2) = "do" And _
            InStr("ct", Mid(Extension, 3, 1)) > 0 And _
            InStr("mx", Right(Extension, 1)) > 0 Then
                Set p_LocalVersion = New Dictionary
                Dim objDoc As Document
                Set objDoc = Documents.Open _
                    (FileName:=FullLocalPath, ReadOnly:=True, Visible:=False)
                p_LocalVersion.Item("major") = _
                    objDoc.CustomDocumentProperties("Major")
                p_LocalVersion.Item("minor") = _
                    objDoc.CustomDocumentProperties("Minor")
                p_LocalVersion.Item("patch") = _
                    objDoc.CustomDocumentProperties("Patch")
        Else
            Set p_LocalVersion = ConfigVersion
        End If
    End If
    Set LocalVersion = p_LocalVersion
LocalVersionGetFinish:
    Exit Property
    
LocalVersionGetError:
    Err.Source = Err.Source & strModuleName & "LocalVersion.Get"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume LocalVersionGetFinish
    End If
End Property

' ===== FullLogPath ===========================================================
' Full path to the log file, including file name and extension. Read only.

Private Property Get FullLogPath() As String
    On Error GoTo FullLogPathGetError
    If p_FullLogPath = vbNullString Then
        p_FullLogPath = Paths_.LocalPath(GroupName:="logFiles") & _
            Application.PathSeparator & Name & "_updates.log"
    End If
    FullLogPath = p_FullLogPath
FullLogPathGetFinish:
    Exit Property

FullLogPathGetError:
    Err.Source = Err.Source & strModuleName & "FullLogPath.Get"
    If SharedMacros_.ErrorChecker(Err, p_FullLogPath) = False Then
        Resume
    Else
        Resume FullLogPathGetFinish
    End If
End Property

' ===== FullTempPath ==========================================================
' Full path to the file in the temp dir. NOT verifying if it's there, because
' we need to reference this before we download it. Read only.

Private Property Get FullTempPath() As String
    On Error GoTo FullTempPathGetError
    If p_FullTempPath = vbNullString Then
        p_FullTempPath = Paths_.TempPath & Application.PathSeparator & FullName
        ' Check if final dir is read-only
        If SharedMacros_.IsReadOnly(p_FullTempPath) = True Then
            Err.Raise MacError.err_TempReadOnly
        End If
    End If
    FullTempPath = p_FullTempPath
FullTempPathGetFinish:
    Exit Property
    
FullTempPathGetError:
    Err.Source = Err.Source & strModuleName & "FullTempPath.Get"
    If SharedMacros_.ErrorChecker(Err, p_FullTempPath) = False Then
        Resume
    Else
        Resume FullTempPathGetFinish
    End If
End Property


' *****************************************************************************
'           PUBLIC METHODS
' *****************************************************************************

' ===== ClearFile =============================================================
' Resets private variables for most properties, in case we, say, update the
' config file during runtime and want to read the new info. Don't want to use
' Property Let at this point because most are read-only, but if we end up
' needing finer control we can add that instead.

Public Sub ClearFile()
    On Error GoTo ClearFileError

    p_Extension = vbNullString
    p_FullLogPath = vbNullString
    p_FullLocalPath = vbNullString
    p_FullName = vbNullString
    p_FullTempPath = vbNullString
    p_Name = vbNullString
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
    
ClearFileFinish:
    Exit Sub
    
ClearFileError:
    Err.Source = Err.Source & strModuleName & "ClearFile"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume ClearFileFinish
    End If
End Sub

' ===== AssignFile ============================================================
' To access all of the properties of the object, both GroupName & SpecificFile
' properties must be set. To make this easier, this method does both at once.

Public Sub AssignFile(strGroupName As String, strSpecificFileType As String)
    On Error GoTo AssignFileError

    GroupName = strGroupName
    SpecificFile = strSpecificFileType

AssignFileFinish:
    Exit Sub

AssignFileError:
    Err.Source = Err.Source & strModuleName & "AssignFile"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume AssignFileFinish
    End If
End Sub

' ===== DeleteFile ============================================================
' Deletes the file in question, as long as it's not currently running this code
' and KillAll() checks if it exists so we don't have to here.

Public Sub DeleteFile()
    On Error GoTo DeleteFileError
    If Running = False Then
        SharedMacros_.KillAll FullLocalPath
    Else
        Err.Raise MacError.err_DeleteThisDoc
    End If
DeleteFileFinish:
    Exit Sub

DeleteFileError:
    Err.Source = Err.Source & strModuleName & "DeleteFile"
    If SharedMacros_.ErrorChecker(Err, FullName) = False Then
        Resume
    Else
        Resume DeleteFileFinish
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
    On Error GoTo NotifyCompleteError
    MsgBox "All downloads are complete. Yay!"
NotifyCompleteFinish:
    Exit Sub

NotifyCompleteError:
    Err.Source = Err.Source & strModuleName & "NotifyComplete"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume NotifyCompleteFinish
    End If
End Sub

' ===== Download ==============================================================
' Downloads the file, which must be posted at the URL returned by the
' DownloadURL property. WinHttp lib not available for Mac, so that uses curl.

' Returns False if there was an error, and it notifies the user of the error in
' the function -- but should probably move this to a global error handler?

Public Function Download() As Boolean
    ' DownloadError section triggers ErrorChecker function. If ErrorChecker
    ' returns True, user is notified, Download = False, & this function exits.
    ' ErrorChecker function writes error info to log, so don't need to here.
    ' Custom errors introduced through Err.Raise must be added to MacError
    ' enum AND to ErrorChecker Select statement.
    On Error GoTo DownloadError
    ' Mac/PC variables
    Dim httpStatus As Long
    Dim strLogMsg As String
    Dim strDownloadURL As String
    
    ' Mac variables
    Dim strBashTmp As String
    
    ' PC variables
    Dim httpRequest As Object
    Dim oStream As Object
    
    ' ----- START -------------------------------------------------------------
    strLogMsg = "Beginning download procedure."
    SharedMacros_.WriteToLog strLogMsg, FullLogPath
    
    ' ----- DELETE TEMP -------------------------------------------------------
    ' Delete the file in the tmp dir if it exists (KillAll checks if exists)
    strLogMsg = "Temp dir clear of file."
    If SharedMacros_.KillAll(FullTempPath) = False Then
        ' delete failed, return an error
        Err.Raise MacError.err_TempDeleteFail
    Else
        SharedMacros_.WriteToLog strLogMsg, FullLogPath
    End If
    
    ' ----- ARE WE CONNECTED TO THE INTERNET? ---------------------------------
    strLogMsg = "Internet connection confirmed."
    #If Mac Then
        ' FullTempPath returns OSX-style path, with colons. Bash shell commands
        ' require forward slashes, and no root volume name.
        strBashTmp = Replace(Right(FullTempPath, Len(FullTempPath) - (InStr _
            (FullTempPath, ":") - 1)), ":", "/")
        If ShellAndWaitMac("ping -o google.com &> /dev/null ; echo $?") <> 0 _
            Then   'can't connect to internet
            Err.Raise MacError.err_NoInternet
        Else
            SharedMacros_.WriteToLog strLogMsg, FullLogPath
        End If
    #Else
        Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
        httpRequest.Open "http://www.google.com"
        httpRequest.Send
        If httpRequest.Status <> 200 Then
            Err.Raise MacError.err_NoInternet
        Else
            SharedMacros_.WriteToLog strLogMsg, FullLogPath
        End If
    #End If
    
    ' ----- CHECK HTTP STATUS OF FILE -----------------------------------------
    strLogMsg = "File is available at URL."
    strDownloadURL = DownloadURL & "/" & Me.FullName
    #If Mac Then
        httpStatus = ShellAndWaitMac("curl -s -o /dev/null -w '%{http_" & _
            "code}' " & strDownloadURL)
    #Else
        httpRequest.Open "HEAD", strDownloadURL
        httpRequest.Send
        httpStatus = httpRequest.Status
    #End If
    
    If httpStatus = 404 Then
        Err.Raise MacError.err_Http404
    ElseIf httpStatus <> 200 Then
        ' Pass httpStatus as Source, status description as Description
        Err.Raise MacError.err_BadHttpStatus, , httpStatus
    Else
        SharedMacros_.WriteToLog strLogMsg, FullLogPath
    End If

    ' ----- DOWNLOAD FILE TO TEMP DIRECTORY -----------------------------------
    strLogMsg = "File download to temp dir successful."
    #If Mac Then
        ShellAndWaitMac ("curl -o " & strBashTmp & " " & strDownloadURL)
    #Else
        httpRequest.Open "GET", strDownloadURL
        httpRequest.Send
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write httpRequest.ResponseBody
        oStream.SaveToFile FullTempPath, 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
        Set oStream = Nothing
        Set httpRequest = Nothing
    #End If

    ' Check if download was successful
    If IsItThere(FullTempPath) = False Then
        Err.Raise MacError.err_DownloadFail
    Else
        SharedMacros_.WriteToLog strLogMsg, FullLogPath
    End If

    ' ----- CLEAR FILE FROM FINAL DIR ------------------------------------------
    strLogMsg = "Final dir clear of previous file."

    ' Check if there is a file in the final dir already
    If IsItThere(Me.FullLocalPath) = False Then
        SharedMacros_.WriteToLog "No previous file in final dir."
    Else
        ' KillAll() checks if file is open and notifies user.
        If SharedMacros_.KillAll(Me.FullLocalPath) = False Then
            Err.Raise MacError.err_LocalDeleteFail
        Else
            SharedMacros_.WriteToLog strLogMsg, FullLogPath
        End If
    End If

    ' ----- MOVE TMP FILE TO FINAL DIR ----------------------------------------
    strLogMsg = "File successfully moved to final dir."
    ' Mac 2011 can't handle file paths > 32 char
    If SharedMacros_.IsOldMac = True Then
        Dim strPosixTemp As String
        Dim strPosixLocal As String
        strPosixTemp = MacScript("return quoted form of posix path of " & _
            FullTempPath)
        strPosixLocal = MacScript("return quoted form of posix path of " & _
            Me.FullLocalPath)
        strMoveCmd = "mv " & strPosixTemp & " " & strPosixLocal
        SharedMacros_.ShellAndWaitMac (strMoveCmd)
    Else
        Name FullTempPath As Me.FullLocalPath
    End If
    
    ' Check if it worked
    If IsItThere(Me.FullLocalPath) = False Then
        Err.Raise MacError.err_LocalCopyFail
    Else
        Download = True
        SharedMacros_.WriteToLog strLogMsg, FullLogPath
    End If
    
    ' ----- SET MAC TEMPLATE ATTRIBUTES ---------------------------------------
    ' Mac won't load macros from a template downloaded from the internet at all
    ' Need to send these commands for it to work, see Confluence docs for info.
    ' Do NOT use open/save-as, this removes customUI.xml, which creates Mac
    ' Tools toolbar later.
    #If Mac Then
        If Extension = "dotm" Then
            Dim strCommand As String
            strCommand = "xattr -wx com.apple.FinderInfo \" & Chr(34) & _
                "57 58 54 4D 4D 53 57 44 00 10 00 00 00 00 00 00 00 00 00 00" _
                & " 00 00 00 00 00 00 00 00 00 00 00 00\" & Chr(34) & Chr(32) _
                & Chr(34) & " & " & "quoted form of POSIX path of " & Chr(34) _
                & Me.FullLocalPath & Chr(34)
                SharedMacros_.ShellAndWaitMac (strCommand)
        End If
    #End If

    'Cleanup: Get rid of temp file if downloaded correctly
    If KillAll(FullTempPath) = False Then
        Err.Raise MacError.err_TempDeleteFail
    End If
DownloadFinish:
    Exit Function
    
DownloadError:
    Err.Source = Err.Source & strModuleName & "Download"
    Dim strDLerror As String
    Select Case Err.Number
        Case MacError.err_TempDeleteFail
            strDLerror = FullTempPath
        Case MacError.err_LocalCopyFail, MacError.err_LocalReadOnly
            strDLerror = FullLocalPath
        Case MacError.err_Http404, MacError.err_DownloadFail, _
            MacError.err_BadHttpStatus
            strDLerror = strHttpStatus & ": " & strDownloadURL
    If SharedMacros_.ErrorChecker(Err, strDLerror) = False Then
        Resume
    Else
        Download = False
        Resume DownloadFinish
    End If
End Function


' *****************************************************************************
'           PRIVATE METHODS
' *****************************************************************************

' ===== StringVersion =========================================================
' returns full version number as a string. Takes an argument so we can use it
' to return either the Config or Local version. Note "patch" may be Null.
'
' May need to make public, in which case should probably be a property? Want to
' keep config and local version dictionaries private, so need another way
' indicate which version. Maybe two properties?

Private Function StringVersion(dictVersion As Dictionary) As String
    On Error GoTo StringVersionError
    Dim strVersionReturn As String
    strVersionReturn = CStr(dictVersion("major")) & "." & _
        CStr(dictVersion("minor"))

    If CStr(dictVersion("patch")) <> vbNullString Then
        strVersionReturn = strVersionReturn & "." & CStr(dictVersion("patch"))
    End If

    StringVersion = strVersionReturn
StringVersionFinish:
    Exit Function
    
StringVersionError:
    Err.Source = Err.Source & strModuleName & "StringVersion"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume StringVersionFinish
    End If
End Function



' *****************************************************************************
'       CLASS EVENTS
' *****************************************************************************

'' ===== Class_Initialize ======================================================
'' Might need this at some point?
'
'Private Sub Class_Initialize()
'
'End Sub

' ===== Class_Terminate =======================================================
' *Object* properties may persist after the calling procedure is ended, even if
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


