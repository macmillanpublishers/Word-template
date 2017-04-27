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
  Else
    DebugPrint "InstallType already set"
  End If
  
  DebugPrint "Current InstallType: " & p_strInstallType
  InstallType = p_strInstallType
End Property

' ===== DownloadBranch ========================================================
' Name of Git branch files should be downloaded from. Branch must be on Github.

' PARAMS
' None

' RETURNS
' String

' NOTES
' Read-only

Public Property Get DownloadBranch() As String
  DownloadBranch = "sectionstart"
End Property
