VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Shared_MacFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ========================================================================
' Class for managing the files that constitute the Macmillan Tools suite

' DEPENDENCIES:
' * config.json file with specific structure, containing info about files,
'   download source, local directory structure
' * https://github.com/VBA-tools/VBA-Dictionary
' * https://github.com/VBA-tools/VBA-Dictionary
' ========================================================================

' ************************************************************************
'           CLASS ENUMERATIONS
' ************************************************************************

    ' This enum should exactly match the first-level keys in the "files"
    ' object in the config JSON file. Is there a way we can pull this
    ' directly from the JSON? Maybe with a DevUtilities macro?
    
Public Enum MacFileGroup
    config = 1
    macroTemplates = 2
    styleTemplates = 4
    assets = 8
End Enum

' ------------------------------------------------------------------------
    ' This enum should exactly match the second-level keys in the "files"
    ' object in the config JSON file. Type should be unique among all files,
    ' so not every Group + Type combo will match an actual file.

Public Enum MacFileGroup
    configFiles = 2 ^ 0     '1
    macroFiles = 2 ^ 1      ' 2
    styleFiles = 2 ^ 2      ' 4
    assetFiles = 2 ^ 3      ' 8
End Enum
' ************************************************************************
'           CLASS CONSTANTS
' ************************************************************************

Private p_FullJSON As Dictionary
Private p_GroupName As MacFileGroup

' ************************************************************************
'           PROPERTIES
' ************************************************************************

    ' Takes full local path to config file as input
    ' Writes it to CustomDocumentProperty for document to read later
    
Private Property Let FullConfigPath(strFullConfigPath As String)
    ThisDocument.CustomDocumentProperties("FullConfigPath").Value = _
        strFullConfigPath
End Property

Private Property Get FullConfigPath() As String
    FullConfigPath = ThisDocument.CustomDocumentProperties("FullConfigPath")
End Property

' -----------------------------------------------------------------------------
    ' Dictionary object of all the data in the JSON config file. We want to
    ' read the whole thing at Initialize so if we need to delete the config
    ' (e.g., to replace with a newer version), this instance can still run.
    ' Read-only (for now) since we'll be managing the JSON externally
    
Private Property Get FullJSON() As Dictionary
    ' Only read JSON if we haven't yet
    If p_FullJSON Is Nothing Then
        Dim strJSON As String
        strJSON = SharedMacros.ReadTextFile(Path:=Me.FullConfigPath, _
            FirstLineOnly:=False)
        Set p_FullJSON = JsonConverter.ParseJson(json_String:=strJSON)
    End If
    FullJSON = p_FullJSON
End Property

' -----------------------------------------------------------------------------
    ' Group here is in reference to the first level of keys in the "file"
    ' object in the JSON file, grouping files with similar purposes together.
    ' The MacFileGroup enum should exactly match those keys.
    ' Do we need to create a "Group" level Dictionary object here?

Private Property Let GroupName(GroupName_value As MacFileGroup)
    p_GroupName = GroupName_value
End Property
Private Property Get GroupName() As MacFileGroup
    GroupName = p_GroupName
End Property

' -----------------------------------------------------------------------------
    ' name of file as a string, WITHOUT extension. We keep name and extension
    ' separate because some things (like log file names) require just the
    ' name w/o the extension.

Private Property Get Name() As String
    
End Property
