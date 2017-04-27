VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WT_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private p_strInstallType As String

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
