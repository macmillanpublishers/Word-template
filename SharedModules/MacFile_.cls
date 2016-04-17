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
    installFiles = startupFiles + macroFiles + styleFiles + assetFiles
    allFiles = configFiles + noConfig
End Enum


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
