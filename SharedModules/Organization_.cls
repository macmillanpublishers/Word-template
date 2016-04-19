VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Organization_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' By Erica Warren - erica.warren@macmillan.com
'
' STATIC CLASS to manage the organization info related to the Macmillan Word
' template. A default instance is always available, so you do not need to Dim
' or Set the object. Just Organization_.HelpEmail and you're good to go.
'
' DEPENDENCIES
' Reads path info from a config JSON file, which must have a top-level
' "organization" object. The full local path of the config file itself must be
' saved as a CustomDocumentProperty in the same template as this module.
' Also requires VBA-Json and VBA-Dictionary, available here:
' https://github.com/VBA-tools
' =============================================================================

' *****************************************************************************
'       PUBLIC PROPERTIES
' *****************************************************************************

' ===== ShortName =============================================================
' Organization name for referencing in code: all lowercase, only alphanumeric
' characters or hyphens.

 Public Property Get ShortName() As String
    ShortName = Me.OrgJSON("shortName")
 End Property

' ===== LongName ==============================================================
' Full name or organization, for like notifications and such.

Public Property Get LongName() As String
    LongName = Me.OrgJSON("longName")
End Property

' ===== HelpEmail =============================================================
' Email address for users to send queries to.

Public Property Get HelpEmail() As String
    HelpEmail = Me.OrgJSON("helpEmail")
End Property


' *****************************************************************************
'       PRIVATE PROPERTIES
' *****************************************************************************

' ===== OrgJSON ===============================================================
' Dictionary object of the "organization" info in the config JSON file.

Private Property Get OrgJSON() As Dictionary
    Set OrgJSON = Paths_.FullJSON("organization")
End Property
