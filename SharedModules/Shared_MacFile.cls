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
' Class for managing the files that constitute the Macmillan tools suite

' DEPENDENCIES:
' config.json file with specific structure, containing info about files
' and download source
' ========================================================================

' ************************************************************************
'           CLASS CONSTANTS
' ************************************************************************

Private p_FullJSON As Dictionary

' ************************************************************************
'           PROPERTIES
' ************************************************************************

Private Property Let FullConfigPath(value As String)
    ' Takes full local path to config file as input
    ' Writes it to CustomDocumentProperty for document to read later
    ThisDocument.CustomDocumentProperties("FullConfigPath").value = value
End Property

Private Property Get FullConfigPath() As String
    ' Read path to config file from CustomDocumentProperty
    ' Doc has to be open to be running this code, so we don't have to open
    ' it first
    
    FullConfigPath = ThisDocument.CustomDocumentProperties("FullConfigPath")
End Property

Private Property Let FullJSON(fullJSON_value As String)

End Property
