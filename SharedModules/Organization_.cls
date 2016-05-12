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
'       PRIVATE VARIABLES
' *****************************************************************************
Private p_ShortName As String
Private p_LongName As String
Private p_HelpEmail As String
Private p_OrgJSON As Dictionary

' For error handling
Private Const strOrg As String = ".Organization_."


' *****************************************************************************
'       PUBLIC PROPERTIES
' *****************************************************************************

' ===== ShortName =============================================================
' Organization name for referencing in code: all lowercase, only alphanumeric
' characters or hyphens.

 Public Property Get ShortName() As String
    On Error GoTo ShortNameError
    If p_ShortName = vbNullString Then
        p_ShortName = Me.OrgJSON("shortName")
    End If
    ShortName = p_ShortName
ShortNameGetFinish:
    Exit Property

ShortNameError:
    Err.Source = Err.Source & strOrg & "ShortName.Get"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume ShortNameGetFinish
    End If
 End Property

' ===== LongName ==============================================================
' Full name or organization, for like notifications and such.

Public Property Get LongName() As String
    On Error GoTo LongNameGetError
    If p_LongName = vbNullString Then
        p_LongName = Me.OrgJSON("longName")
    End If
    LongName = p_LongName
LongNameGetFinish:
    Exit Property

LongNameGetError:
    Err.Source = Err.Source & strOrg & "LongName.Get"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume LongNameGetFinish
    End If
End Property

' ===== HelpEmail =============================================================
' Email address for users to send queries to.

Public Property Get HelpEmail() As String
    On Error GoTo HelpEmailGetError
    If p_HelpEmail = vbNullString Then
        p_HelpEmail = OrgJSON("helpEmail")
    End If
    HelpEmail = p_HelpEmail
HelpEmailGetFinish:
    Exit Property

HelpEmailGetError:
    Err.Source = Err.Source & strOrg & "HelpEmail.Get"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume HelpEmailGetFinish
    End If
End Property


' *****************************************************************************
'       PRIVATE PROPERTIES
' *****************************************************************************

' ===== OrgJSON ===============================================================
' Dictionary object of the "organization" info in the config JSON file.

Private Property Get OrgJSON() As Dictionary
    On Error GoTo OrgJsonGetError
    If p_OrgJSON Is Nothing Then
        Set p_OrgJSON = Paths_.FullJSON("organization")
    End If
    Set OrgJSON = p_OrgJSON
OrgJsonGetFinish:
    Exit Property

OrgJsonGetError:
    Err.Source = Err.Source & strOrg & "OrgJSON.Get"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume OrgJsonGetFinish:
    End If
End Property
