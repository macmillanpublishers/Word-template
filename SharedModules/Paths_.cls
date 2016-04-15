VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Paths_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' =============================================================================
' By Erica Warren - erica.warren@macmillan.com
'
' STATIC CLASS to manage properties and methods related to the directory
' structure of the Macmillan Word template. A default instance is always
' available, so you do not need to Dim or Set the object. Just
' Paths_.HomeDir and you're good to go.
'
' DEPENDENCIES
' Reads path info from a config JSON file, which must have a top-level "paths"
' object. The full local path of the config file itself must be saved as a
' CustomDocumentProperty in the same template as this module.
' Also requires VBA-Json and VBA-Dictionary, available here:
' https://github.com/VBA-tools
' =============================================================================

'
