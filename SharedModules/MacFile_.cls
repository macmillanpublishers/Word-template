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


' *****************************************************************************
'           ERROR HANDLING STUFF
' *****************************************************************************
' Technically class methods (which are listed farther down) but it's easier to
' maintain close to the error number enum, which has to appear above other
' procedures. May move to its own class one day.

' ===== MacFileErrors =========================================================
' Each error we raise anywhere in this class should be included here. Note that
' error numbers 0 - 512 are reserved for system errors. Include property or
' method name in enum just to be clear for debugging.

Private Enum MacFileError
    MacFileErrGeneral = 20000       ' Not sure if we'll need, but here it is.
    GroupNameErr1 = 20001
    GroupNameErr2 = 20002
    SpecificFileErr1 = 20003
    SpecificFileErr2 = 20004
    DeleteFileErr1 = 20005
    DownloadErr1 = 20006
    DownloadErr2 = 20007
    DownloadErr3 = 20008
    DownloadErr4 = 20009
    DownloadErr5 = 20010
    DownloadErr6 = 20011
    DownloadErr7 = 20012
    DownloadErr8 = 20013
    DownloadErr9 = 20014
    FullLocalPathErr1 = 20015
End Enum

' ===== ErrorChecker ==========================================================
' Send all errors here, to make tracking/maintaining error handling easier.
' Be sure to add the error to the MacFileErrors enum.
'
' USE: ErrorChecker() returns False if error is handled; procedure can continue
' Each member procedure in the class should have the following:
'''
' On Error GoTo MemberName[Let | Set | Get]Error
'   <...code...>
' MemberName[Let | Set | Get]Finish:
'   <...any cleanup code...>
'   On Error GoTo 0
'   Exit [Property | Sub | Function]
' MemberName[Let | Set | Get]Error:
'    If ErrorChecker(Err) = False Then
'        Resume
'    Else
'        Resume MemberName[Let | Set | Get]Finish
'    End If
' End [Property | Sub | Function]

Private Function ErrorChecker(ByRef objError As ErrObject) As Boolean
    Dim strDescription As String
    Dim strErrMessage As String
    Dim blnNotifyUser As Boolean
    Dim strHelpContact As String

    ' ----- Set defaults ------------------------------------------------------
    ' Only need to change in Select statement below if want to set either to
    ' False.
    blnNotifyUser = True
    ErrorChecker = True
    strHelpContact = "Email " & Organization_.HelpEmail & " if you need help."

    ' ----- Check for errors --------------------------------------------------
    ' Make sure we actually have an error, cuz you never know.
    If objError.Number > 0 Then

    ' ----- Handle specific errors --------------------------------------------
    ' Each case should be a MacFileError enum. Even if we aren't doing anything
    ' to fix the error, still enter the following:
    '   Err.Source "MacFile_.<MethodName>"
    '   strDescription = "Description of the error for the log."
    '   strErrMessage = "Message for the user if we're notifying them."
        Select Case objError.Number
            Case Is < 513
                ' Errors 0 to 512 are system errors
                strDescription = objError.Description
                strErrMessage = "Something unexpected happened. Please click" _
                    & " OK to exit." & vbNewLine & vbNewLine & strHelpContact
            Case MacFileError.GroupNameErr1
                objError.Source = "MacFile_.GroupName " & Me.FullName
                strDescription = "Invalid value for GroupName property."
                strErrMessage = "The value you've entered for the GroupName " _
                    & "property is not valid. Make sure you only use file " & _
                    "groups that are in the config.json file."
            Case MacFileError.GroupNameErr2
                objError.Source = "MacFile_.GroupName " & Me.FullName
                strDescription = "GroupName property has not been Let."
                strErrMessage = "You can't Get the GroupName property before" _
                    & " it has been Let. Try the MacFile_.AssignFile method " _
                    & "to create a new object in this class."
            Case MacFileError.SpecificFileErr1
                objError.Source = "MacFile_.SpecificFile " & Me.FullName
                strDescription = "Invalid value for SpecificFile property."
                strErrMessage = "The you've entered an invalid value for the" _
                    & " SpecificFile property. Make sure you only use " & _
                    "specific file types that are in the config.json file."
            Case MacFileError.SpecificFileErr2
                objError.Source = "MacFile_.SpecificFile " & Me.FullName
                strDescription = "SpecificFile property has not been Let."
                strErrMessage = "You can't Get the SpecificFile property " & _
                    "before it has been Let. Try the MacFile_.AssignFile " & _
                    "method to create a new object in this class."
            Case MacFileError.DeleteFileErr1
                objError.Source = "MacFile_.DeleteFile " & Me.FullName
                strDescription = "File is currently executing code."
                strErrMessage = "The file you are trying to delete is " & _
                    "currently executing macro code."
            Case MacFileError.DownloadErr1
                objError.Source = "MacFile_.Download " & Me.FullName
                strDescription = "Failed to delete the previous file in the " _
                    & "temp directory. Download aborted."
                strErrMessage = "We weren't able to download the file " & _
                    Me.FullName & " because of a stubborn temp file. " & _
                    vbNewLine & vbNewLine & strHelpContact
            Case MacFileError.DownloadErr2
                objError.Source = "MacFile_.Download " & Me.FullName
                strDescription = "No network connection. Download aborted."
                strErrMessage = "We weren't able to download the file " & _
                    Me.FullName & " because we can't connect to the " & _
                        "internet. Check your network connection and try " & _
                        "again." & vbNewLine & vbNewLine & strHelpContact
            Case MacFileError.DownloadErr3
                objError.Source = "MacFile_.Download " & Me.FullName
                strDescription = "File HTTP status 404. Check if DownloadURL" _
                    & " is correct, and file is posted."
                strErrMessage = "Could not download " & Me.FullName & _
                    " from the internet. " & strHelpContact
            Case MacFileError.DownloadErr4
                strDescription = "File HTTP status " & objError.Source & ": " _
                    & objError.Description & ". Download aborted."
                strErrMessage = "There is some problem with the file you are" _
                    & " trying to download: " & Me.FullName & vbNewLine & _
                    objError.Source & ": " & objError.Description & vbNewLine _
                    & vbNewLine & strHelpContact
                ' Need to get Source as passed in object first, so do this last
                objError.Source = "MacFile_.Download " & Me.FullName
            Case MacFileError.DownloadErr5
                objError.Source = "MacFile_.Download " & Me.FullName
                strDescription = "File download failed."
                strErrMessage = "Download of " & Me.FullName & " failed. " & _
                    strHelpContact
            Case MacFileError.DownloadErr6
                objError.Source = "MacFile_.Download " & Me.FullName
                strDescription = "File in final install location could not " _
                    & "be deleted. If it was because the file was open, the " _
                    & "user was notified."
                blnNotifyUser = False
            Case MacFileError.FullLocalPathErr1
                objError.Source = "MacFile_.FullLocalPath " & Me.FullName
                strDescription = "Final dir for file is read-only."
                strErrMessage = "The folder you are trying to access, " & _
                    Me.FullLocalPath & ", is read-only. " & strHelpContact
            Case Else
                objError.Source = "MacFile_.? " & Me.FullName
                strDescription = "Undocumented error number - " & _
                    objError.Number & ": " & objError.Description
                strErrMessage = "Not sure what's going on here." & strHelpEmail
        End Select

    Else
        objError.Source = "MacFile_ " & Me.FullName
        strDescription = "Everything's A-OK. Why are you even reading this?"
        blnNotifyUser = False
        ErrorChecker = False
    End If

    WriteToLog objError.Source & " - " & objError.Number & ": " & _
        strDescription

    If blnNotifyUser = True Then
        MsgBox Prompt:=strErrMessage, Buttons:=vbExclamation, Title:= _
            "Macmillan Tools Error: " & objError.Number & " " & Me.FullName
    End If

    objError.Clear
End Function


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
        Err.Raise MacFileError.GroupNameErr1
    End If
GroupNameLetFinish:
    On Error GoTo 0
    Exit Property

GroupNameLetError:
    If ErrorChecker(Err) = False Then
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
        Err.Raise MacFileError.GroupNameErr2
    End If
GroupNameGetFinish:
    On Error GoTo 0
    Exit Property

GroupNameGetError:
    If ErrorChecker(Err) = False Then
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
        Err.Raise MacFileError.SpecificFileErr1
    End If
SpecificFileLetFinish:
    On Error GoTo 0
    Exit Property

SpecificFileLetError:
    If ErrorChecker(Err) = False Then
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
        Err.Raise MacFileError.SpecificFileErr2
    End If
SpecificFileGetFinish:
    On Error GoTo 0
    Exit Property

SpecificFileGetError:
    If ErrorChecker(Err) = False Then
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
    On Error GoTo 0
    Exit Property
    
FullNameGetError:
    If ErrorChecker(Err) = False Then
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
            Err.Raise FullLocalPathErr1
        End If
    End If
    FullLocalPath = p_FullLocalPath
FullLocalPathGetFinish:
    On Error GoTo 0
    Exit Property
    
FullLocalPathGetError:
    If ErrorChecker(Err) = False Then
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
    On Error GoTo 0
    Exit Property
    
ExistsGetError:
    If ErrorChecker(Err) = False Then
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
    If ThisDocument.Name = FullName Then
        Running = True
    Else
        Running = False
    End If
RunningGetFinish:
    On Error GoTo 0
    Exit Property

RunningGetError:
    If ErrorChecker(Err) = False Then
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
    On Error GoTo 0
    Exit Property
    
UpToDateGetError:
    If ErrorChecker(Err) = False Then
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
    On Error GoTo 0
    Exit Property

GroupJSONGetError:
    If ErrorChecker(Err) = False Then
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
    On Error GoTo 0
    Exit Property
    
FileJSONGetError:
    If ErrorChecker(Err) = False Then
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
    On Error GoTo 0
    Exit Property

NameGetError:
    If ErrorChecker(Err) = False Then
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
    On Error GoTo 0
    Exit Property
    
ExtensionGetError:
    If ErrorChecker(Err) = False Then
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
    On Error GoTo 0
    Exit Property
    
ConfigVersionGetError:
    If ErrorChecker(Err) = False Then
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
    On Error GoTo 0
    Exit Property
    
LocalVersionGetError:
    If ErrorChecker(Err) = False Then
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
    On Error GoTo 0
    Exit Property

FullLogPathGetError:
    If ErrorChecker(Err) = False Then
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
    End If
    FullTempPath = p_FullTempPath
FullTempPathGetFinish:
    On Error GoTo 0
    Exit Property
    
FullTempPathGetError:
    If ErrorChecker(Err) = False Then
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
    
ClearFileFinish:
    On Error GoTo 0
    Exit Sub
    
ClearFileError:
    If ErrorChecker(Err) = False Then
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
    On Error GoTo 0
    Exit Sub

AssignFileError:
    If ErrorChecker(Err) = False Then
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
        Err.Raise MacFileError.DeleteFileErr1
    End If
DeleteFileFinish:
    On Error GoTo 0
    Exit Sub

DeleteFileError:
    If ErrorChecker(Err) = False Then
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
    On Error GoTo 0
    Exit Sub

NotifyCompleteError:
    If ErrorChecker(Err) = False Then
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
    ' Custom errors introduced through Err.Raise must be added to MacFileError
    ' enum AND to ErrorChecker Select statement.
    On Error GoTo DownloadError
    ' Mac/PC variables
    Dim httpStatus As Long
    Dim strLogMsg As String
    
    ' Mac variables
    Dim strBashTmp As String
    
    ' PC variables
    Dim httpRequest As Object
    Dim oStream As Object
    
    ' ----- START -------------------------------------------------------------
    WriteToLog "Beginning download procedure."
    
    ' ----- DELETE TEMP -------------------------------------------------------
    ' Delete the file in the tmp dir if it exists (KillAll checks if exists)
    strLogMsg = "Temp dir clear of file."
    If SharedMacros_.KillAll(FullTempPath) = False Then
        ' delete failed, return an error
        Err.Raise MacFileError.DownloadErr1
    Else
        WriteToLog strLogMsg
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
            Err.Raise MacFileError.DownloadErr2
        Else
            WriteToLog strLogMsg
        End If
    #Else
        Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
        httpRequest.Open "http://www.google.com"
        httpRequest.Send
        If httpRequest.Status <> 200 Then
            Err.Raise MacFileError.DownloadErr2
        Else
            WriteToLog strLogMsg
        End If
    #End If
    
    ' ----- CHECK HTTP STATUS OF FILE -----------------------------------------
    strLogMsg = "File is available at URL."
    #If Mac Then
        httpStatus = ShellAndWaitMac("curl -s -o /dev/null -w '%{http_" & _
            "code}' " & DownloadURL)
    #Else
        httpRequest.Open "HEAD", DownloadURL
        httpRequest.Send
        httpStatus = httpRequest.Status
    #End If
    
    If httpStatus = 404 Then
        Err.Raise MacFileError.DownloadErr3
    ElseIf httpStatus <> 200 Then
        ' Pass httpStatus as Source, status description as Description
        Err.Raise MacFileError.DownloadErr4, httpStatus, httpRequest.StatusText
    Else
        WriteToLog strLogMsg
    End If

    ' ----- DOWNLOAD FILE TO TEMP DIRECTORY -----------------------------------
    strLogMsg = "File download to temp dir successful."
    #If Mac Then
        ShellAndWaitMac ("curl -o " & strBashTmp & " " & DownloadURL)
    #Else
        httpRequest.Open "GET", DownloadURL
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
        Err.Raise MacFileError.DownloadErr5
    Else
        WriteToLog strLogMsg
    End If

    ' ----- CLEAR FILE FROM FINAL DIR ------------------------------------------
    strLogMsg = "Final dir clear of previous file."

    ' Check if there is a file in the final dir already
    If IsItThere(Me.FullLocalPath) = False Then
        WriteToLog "No previous file in final dir."
    Else
        If SharedMacros_.KillAll(Me.FullLocalPath) = False Then
            Err.Raise MacFileError.DownloadErr6
        Else
            WriteToLog strLogMsg
        End If
    End If

    ' ----- MOVE TMP FILE TO FINAL DIR ----------------------------------------
    strLogMsg = "File successfully moved to final dir."
    ' Mac 2011 can't handle file paths > 32 char
    #If Mac And Application.Version < 16 Then
        Dim strPosixTemp As String
        Dim strPosixLocal As String
        strPosixTemp = MacScript("return quoted form of posix path of " & _
            FullTempPath)
        strPosixLocal = MacScript("return quoted form of posix path of " & _
            Me.FullLocalPath)
        strMoveCmd = "mv " & strPosixTemp & " " & strPosixLocal
        SharedMacros_.ShellAndWaitMac (strMoveCmd)
    #Else
        Name FullTempPath As Me.FullLocalPath
    #End If
    
    ' Check if it worked
    If IsItThere(Me.FullLocalPath) = False Then
        Err.Raise MacFileError.DownloadErr7
    Else
        WriteToLog strLogMsg
    End If
    
    ' ----- SET MAC TEMPLATE ATTRIBUTES ---------------------------------------
    ' Mac won't load macros from a template downloaded from the internet at all
    ' Need to send these commands for it to work, see Confluence docs for info.
    ' Do NOT use open/save-as, this removes customUI.xml which creates Mac
    ' Tools toolbar later.
    #If Mac Then
        If Extension = "dotm" Then
            Dim strCommand As String
            strCommand = "do shell script " & Chr(34) & "xattr -wx " & _
                "com.apple.FinderInfo \" & Chr(34) & "57 58 54 4D 4D 53 57" & _
                " 44 00 10 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00" & _
                "00 00 00 00 00 00\" & Chr(34) & Chr(32) & Chr(34) & " & " & _
                "quoted form of POSIX path of " & Chr(34) & strFinalPath & Chr(34)
                'Debug.Print strCommand
                MacScript (strCommand)
        End If
    #End If

    Else
        strLogMsg = Now & " -- old " & FullName & " file not cleared from Final directory."
        WriteToLog LogMessage:=strLogMsg
        strErrMsg = "There was an error installing the Macmillan template." & vbNewLine & _
            "Please close all other Word documents and try again, or contact workflows@macmillan.com."
        MsgBox strErrMsg, vbCritical, "Error 5: Previous version uninstall failed (" & FullName & ")"
        Download = False
        On Error GoTo 0
        Exit Function
    End If

    'If move was successful, yay! Else, :(
    If IsItThere(strFinalPath) = True Then
        strLogMsg = Now & " -- " & FullName & " file successfully saved to final directory."
        WriteToLog LogMessage:=strLogMsg
    Else
        strLogMsg = Now & " -- " & FullName & " file not saved to final directory."
        WriteToLog LogMessage:=strLogMsg
        strErrMsg = "There was an error installing the Macmillan template." & vbNewLine & vbNewLine & _
            "Please cotact workflows@macmillan.com for assistance."
        MsgBox strErrMsg, vbCritical, "Error 6: Installation failed (" & FullName & ")"
        Download = False
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

    Download = True
DownloadFinish:
    On Error GoTo 0
    Exit Function
    
DownloadError:
    If ErrorChecker(Err) = False Then
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
    On Error GoTo 0
    Exit Function
    
StringVersionError:
    If ErrorChecker(Err) = False Then
        Resume
    Else
        Resume StringVersionFinish
    End If
End Function

' ===== WriteToLog ============================================================
' Writes line to log for the file. LogMessage only needs text, timestamp will
' be added in this method. May make public later if we end up using this for
' more than just download logs.

Private Sub WriteToLog(LogMessage As String)
    On Error GoTo WriteToLogError
    Dim strLogMessage As String
    Dim FileNum As Integer

    strLogMessage = Now & " -- " & LogMessage
    FileNum = FreeFile ' next file number

    Open LogFile For Append As #FileNum ' creates the file if it doesn't exist
    Print #FileNum, strLogMessage ' write information to end of the text file
    Close #FileNum ' close the file
WriteToLogFinish:
    On Error GoTo 0
    Exit Sub

WriteToLogError:
    If ErrorChecker(Err) = False Then
        Resume
    Else
        Resume WriteToLogFinish
    End If
End Sub



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


