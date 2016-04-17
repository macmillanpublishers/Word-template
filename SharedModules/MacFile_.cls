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

' ------------------------------------------------------------------------
    ' This enum should exactly match the first-level keys in the "files"
    ' object in the config JSON file. Must be set to powers of two so
    ' we can use bitwise flags to combine groups.

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


' ************************************************************************
'           CLASS CONSTANTS
' ************************************************************************

Private p_GroupName As MacFileGroup


' ************************************************************************
'           PROPERTIES
' ************************************************************************

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
