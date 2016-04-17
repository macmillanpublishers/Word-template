VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MacFile_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' =============================================================================
' Class for managing the files that constitute the Macmillan Tools suite

' DEPENDENCIES:
' * config.json file with specific structure, containing info about files,
'   download source, local directory structure
' * Paths_.cls static class module
' * https://github.com/VBA-tools/VBA-Dictionary
' * https://github.com/VBA-tools/VBA-JSON
' =============================================================================


' *****************************************************************************
'           CLASS ENUMERATIONS
' *****************************************************************************

' ===== MacFileGroup ==========================================================
' This enum should include the first-level keys in the "files" object in
' the config JSON file. Must be set to powers of two so we can use bitwise
' flags to combine groups.

Public Enum MacFileGroup
    configFiles = 2 ^ 0     ' 1
    startupFiles = 2 ^ 1    ' 2
    macroFiles = 2 ^ 2      ' 4
    styleFiles = 2 ^ 3      ' 8
    assetFiles = 2 ^ 4      ' 16
    updateCheck = macroFiles + styleFiles
    wordDocs = startupFiles + macroFiles + styleFiles
    installFiles = startupFiles + macroFiles + styleFiles + assetFiles
    allFiles = configFiles + noConfig
End Enum


' *****************************************************************************
'           CLASS DATA TYPES
' *****************************************************************************
' Data type to manage version numbers, which follow Major.Minor.Patch format as
' described here: http://semver.org/. Version in JSON must be stored as three
' separate integers, Major and Minor must be an integer (can be 0) but Patch
' can be null.

Public Type VersionNumber
    Major As Long
    Minor As Long
    Patch As Long
End Type


' *****************************************************************************
'           CLASS VARIABLES
' *****************************************************************************

Private p_FileGroup As String
Private p_SpecificFile As String


' *****************************************************************************
'           PUBLIC PROPERTIES
' *****************************************************************************

' ===== FileGroup =============================================================
' Group here is in reference to the first level of keys in the "file" object in
' JSON file, grouping files with similar purposes together.

Public Property Let FileGroup(FileGroup_value As String)
    If Paths_.FullJSON("files").Exists(FileGroup_value) Then
        p_FileGroup = FileGroup_value
    Else
        MsgBox "That's not a valid option for this property."
    End If
End Property

Public Property Get FileGroup() As String
    If p_FileGroup <> vbNullString Then
        FileGroup = p_FileGroup
    Else
        MsgBox "You must set the FileGroup property first!"
    End If
End Property

' ===== SpecificFile ==========================================================
' Specific meaning "category" or "role"; one level below FileGroup in the JSON
' file; i.e., what we are calling the file based on what it does in the project
' (not its Name, which can change). Options for specific file are limited to
' the keys of the various objects in the FileGroups, and each must me unique
'
' The SpecificFile property needs to be set before you can access other info.

Public Property Let SpecificFile(SpecificFile_value As String)
    If Me.GroupJSON.Exists(SpecificFile_value) Then
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
    FullName = Me.Name & "." & Me.Extension
End Property

' ===== FullLocalPath =========================================================
' Returns string of the full path to the file, including root and full file
' name with extension.

Public Property Get FullLocalPath() As String
    FullLocalPath = Paths_.LocalPath(GroupName:=Me.FileGroup) & _
        Application.PathSeparator & Me.FullName
End Property

' ===== Exists ================================================================
' Tests whether the files exists at its FullLocalPath

Public Property Get Exists() As Boolean
    Exists = SharedMacros.IsItThere(Path:=Me.FullLocalPath)
End Property

' ===== Running ===============================================================
' Tests whether the file in question is the one that's currently running this
' code, so we don't like delete ourselves or something.

Public Property Get Running() As Boolean
    If ThisDocument.Name = Me.FullName Then
        Running = True
    Else
        Running = False
    End If
End Property


' *****************************************************************************
'           PRIVATE PROPERTIES
' *****************************************************************************

' ===== GroupJSON =============================================================
' Dictionary object of the FileGroup level in the JSON file. Not entirely sure
' we definitely need this here, but then I referenced it in two places so I
' figured why not.

Private Property Get GroupJSON() As Dictionary
    Set GroupJSON = Paths_.FullJSON("files")(Me.FileGroup)
End Property

' ===== FileJSON ==============================================================
' Dictionary object of the info for the specific file, referenced by the
' SpecificFile property.

Private Property Get FileJSON() As Dictionary
    Set FileJSON = Me.GroupJSON(Me.SpecificFile)
End Property

' ===== Name ==================================================================
' Name of file as a string, WITHOUT extension. Keep name and extension separate
' because some things (like log file names) require just the name w/o extension.

Private Property Get Name() As String
    Name = Me.FileJSON("fileName")
End Property

' ===== Extension =============================================================
' File extension as a string, NOT including period.

Private Property Get Extension() As String
    Extension = Me.FileJSON("extension")
End Property

' ===== ConfigVersion ==========================================================
' Version of the file as noted in the config.JSON file. Note that Patch can be
' null but Major and Minor cannot.

Private Property Get ConfigVersion() As VersionNumber
    Dim dictVersion As Dictionary
    Set dictVersion = Me.FileJSON("version")
    ConfigVersion.Major = dictVersion("major")
    ConfigVersion.Minor = dictVersion("minor")
    ConfigVersion.Patch = dictVersion("patch")
End Property

' ===== LocalVersion ==========================================================
' For Word documents/templates: version of the file as noted in the Custom
' Document Properties. Note, stored as three separate properties for Major,
' Minor, and Patch. Patch may be null. Document must be open (though might be a
' way to manage this by unzipping the file and reading XML directly?).

Private Property Get LocalVersion() As VersionNumber
    If Left(Me.Extension, 2) = "do" And _
        InStr("ct", Mid(Me.Extension, 3, 1)) > 0 And _
        InStr("mx", Right(Me.Extension, 1)) > 0 Then
        
        Dim objDoc As Document
        Set objDoc = Documents.Open _
            (FileName:=Me.FullLocalPath, ReadOnly:=True, Visible:=False)
        LocalVersion.Major = objDoc.CustomDocumentProperties("Major")
        LocalVersion.Minor = objDoc.CustomDocumentProperties("Minor")
        LocalVersion.Pathc = objDoc.CustomDocumentProperties("Patch")
    Else
        LocalVersion = Me.ConfigVersion
    End If
End Property

' ===== DownloadURL ===========================================================
' Full URL to download the file, including file name.

Private Property Get DownloadURL() As String
    Dim dict
End Property
