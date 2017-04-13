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
      ' user is "padwoadmin" on both servers
      strLocalUser = Environ("USERNAME")
      If strLocalUser = "padwoadmin" Then ' we're on the server
        p_strInstallType = "server"
      ' If not, just use desktop
      Else
        p_strInstallType = "user"
      End If
    End If
    
    InstallType = p_strInstallType
End Property
