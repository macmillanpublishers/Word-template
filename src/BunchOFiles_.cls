VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BunchOFiles_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' By Erica Warren - erica.warren@macmillan.com
'
' Class to manupulate groups of MacFile_ objects.
'
' DEPENDENCIES
' MacFile_.cls custom class.
' Also requires VBA-Json and VBA-Dictionary, available here:
' https://github.com/VBA-tools
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

