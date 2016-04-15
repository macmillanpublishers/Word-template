VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Shared_MacFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' This enum should exactly match the first-level items in the "files"
' object in the config JSON file
Public Enum GroupType
    config = 1
    macroTemplates = 2
    styleTemplates = 3
    assets = 4
End Enum

' ************************************************************************
'           CLASS CONSTANTS
' ************************************************************************

Private p_FullJSON As Dictionary

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
    ' While this can be a string, should it be an Enum, because the only
    ' acceptable values are those

Private Property Get GroupName() As String

End Property
' -----------------------------------------------------------------------------
    ' name of file as a string, WITHOUT extension. We keep name and extension
    ' separate because some things (like log file names) require just the
    ' name w/o the extension.

Private Property Get Name() As String

End Property
