VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WT_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
'       WT_Settings
' =============================================================================

' ***** USE *******************************************************************
' Static class to manage settings for the Word-template. Can access at any time
' by using it's name directly: WT_Settings.PropertyName

' ***** DEPENDENCIES **********************************************************
' Static class: there is a single default instance that is always available, so
' you don't need to (in fact, can't) create multiple objects from it.
' To set as a static class (only need to do once):
' 1. Export the class module.
' 2. Open the exported file in a text editor.
' 3. On line 8, set Attribute VB_PredeclaredId = True
' 4. Save .cls file.
' 5. Back in VB Editor, remove (and don't save) current class module.
' 6. Import the class module you edited above.

' ***** TO DO *****************************************************************
' [ ] Change all properties to read from config file

' *****************************************************************************
'       DECLARATIONS
' *****************************************************************************
Option Explicit

Private p_strInstallType As String
Private p_Overrides As Dictionary


' *****************************************************************************
'     PROCEDURES
' *****************************************************************************

' ***** PUBLIC PROPERTIES *****************************************************

' ===== InstallType ===========================================================
' Word-template.dotm can be used in multiple settings, notably on the server to
' run bookmaker_validator vs. by end-users (via Macmillan Tools tab). This returns
' the current place the template is installed.

Public Property Get InstallType() As String
  If p_strInstallType = vbNullString Then
    Dim strLocalUser As String
    strLocalUser = Environ("USERNAME")

  ' user is "padwoadmin" on both servers
  ' TODO: change this to read from config file
    Select Case strLocalUser
      Case "padwoadmin"
        p_strInstallType = "server"
      Case "erica.warren"
        p_strInstallType = "debug"
      Case "matthew.retzer"
        p_strInstallType = "debug"
      Case Else
        p_strInstallType = "user"
    End Select
    
    If p_strInstallType = "debug" And Overrides("InstallType").Exists(p_strInstallType) Then
      p_strInstallType = Overrides("InstallType")(p_strInstallType)
    End If
  Else
    DebugPrint "InstallType already set"
  End If
  
  DebugPrint "Current InstallType: " & p_strInstallType
  InstallType = p_strInstallType
End Property

' ===== DownloadBranch ========================================================
' Name of Git branch files should be downloaded from. Branch must be on Github.
' Users always default to "master" but debug can pull from anything.

' PARAMS
' Repo: name of the

' RETURNS
' String

' NOTES
' Read-only

Public Property Get DownloadBranch(Repo As String) As String

  If Me.InstallType = "debug" And Overrides("DownloadBranch").Exists(Repo) Then
    DownloadBranch = Overrides("DownloadBranch")(Repo)
  Else
    DownloadBranch = "master"
  End If
End Property

' =====

' ***** PRIVATE PROPERTIES *****************************************************

' ===== Overrides ==============================================================
' Read JSON file to override certain settings when debugging. Must be in same
' directory as this file.

Private Property Get Overrides() As Dictionary
  If p_Overrides Is Nothing Then
    Dim strFile As String
    strFile = ThisDocument.Path & Application.PathSeparator & "settings_overrides.json"
  
    If Utils.IsItThere(strFile) = True Then
      Set p_Overrides = ClassHelpers.ReadJson(strFile)
    Else
      Set p_Overrides = New Dictionary
    End If
  End If

  Set Overrides = p_Overrides
  
End Property






