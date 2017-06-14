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
Private p_GlobalConfig As Dictionary
Private p_DefaultConfigPath As String
Private p_TmpDir As String
Private p_Region As String
Private p_LocalConfig As Dictionary


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

  ' username is "padwoadmin" on both servers
  ' TODO: change this to read from config file
    Select Case strLocalUser
      Case "padwoadmin"
        p_strInstallType = "server"
      Case Else
        p_strInstallType = "user"
    End Select

  Else
'    DebugPrint "InstallType already set"
  End If
  
'  DebugPrint "Current InstallType: " & p_strInstallType
  InstallType = p_strInstallType
End Property


' ===== DebugOn ===============================================================
' Are we in debugging mode? NOTE: can't use 'Debug' because that's reserved!

Public Property Get DebugOn() As Boolean
  If LocalConfig.Exists("debug") = True Then
    DebugOn = LocalConfig("debug")
  End If
End Property


' ===== StyleDir ==============================================================
' Returns path to MacmillanStyleTemplate directory on local machine. If you pass
' "GtUpdater.dotm" as FileName, it returns the Startup folder.


Public Property Get StyleDir(Optional FileName As String) As String
  Dim strFullPath As String

' GtUpdater.dotm only file to go in Startup
  If FileName = "GtUpdater.dotm" Then
    strFullPath = Application.StartupPath
  Else
    Dim strMacDocs As String
    Dim strStylesName As String
    strStylesName = "MacmillanStyleTemplate"
  
    #If Mac Then
      strMacDocs = MacScript("return (path to documents folder) as string")
      strFullPath = strMacDocs & strStylesName
    #Else
      strFullPath = Environ("APPDATA") & Application.PathSeparator & strStylesName
    #End If
  
  ' Create directory if it doesn't exist yet
    If Utils.IsItThere(strFullPath) = False Then
      MkDir strFullPath
    End If
  End If

  StyleDir = strFullPath

End Property


' ===== Region ================================================================
' What region (us or uk) is user in?

Public Property Get Region() As String
  If p_Region = vbNullString Then
    p_Region = LocalConfig("region")
  End If
  Region = p_Region
End Property





' ***** PRIVATE PROPERTIES *****************************************************

' ===== TmpDir ================================================================
' Returns path to TEMP directory, no trailing path separator

Private Property Get TmpDir() As String
  If p_TmpDir = vbNullString Then
    #If Mac Then
      p_TmpDir = MacScript("path to temporary items as string")
    #Else
      p_TmpDir = Environ("TEMP")
    #End If
  ' Remove trailing path separator, if any
    If Right(p_TmpDir, 1) = Application.PathSeparator Then
      p_TmpDir = Left(strTmpDir, Len(p_TmpDir) - 1)
    End If
  End If
  TmpDir = p_TmpDir
End Property


' ===== DefaultConfigPath =====================================================
' Default path to config files on local machine.

Private Property Get DefaultConfigPath() As String
  If p_DefaultConfigPath = vbNullString Then
    p_DefaultConfigPath = StyleDir & Application.PathSeparator & "config"
    If Utils.IsItThere(p_DefaultConfigPath) = False Then
      MkDir p_DefaultConfigPath
    End If
  End If
    DefaultConfigPath = p_DefaultConfigPath
End Property


' ===== GlobalConfig ==========================================================
' Dictionary of data in global_config.json file. If file isn't present, it is
' downloaded from the URL saved in the template's CustomDocumentProperties.

Private Property Get GlobalConfig() As Dictionary
  If p_GlobalConfig Is Nothing Then
    Dim strFileName As String
    Dim strFullPath As String
    strFileName = "global_config.json"
    strFullPath = DefaultConfigPath & Application.PathSeparator & strFileName

    If Utils.IsItThere(strFullPath) = False Then
      DownloadFromGithub strFullPath
    End If
    Set p_GlobalConfig = ClassHelpers.ReadJson(strFileName)
  End If
  
  Set GlobalConfig = p_GlobalConfig
End Property

' ===== RegionConfig ==========================================================
' Dictionary of data in global_config.json file. If file isn't present, it is
' downloaded from the URL saved in the template's CustomDocumentProperties.

Private Property Get RegionConfig() As Dictionary
  If p_RegionConfig Is Nothing Then
    Dim strFileName As String
    Dim strFullPath As String
    strFileName = Me.Region & "_config.json"
    strFullPath = DefaultConfigPath & Application.PathSeparator & strFileName

    If Utils.IsItThere(strFullPath) = False Then
      DownloadFromGithub strFullPath
    End If
    Set p_RegionConfig = ClassHelpers.ReadJson(strFileName)
  End If
  
  Set RegionConfig = p_RegionConfig
End Property

' ===== LocalConfig ===========================================================
' Local settings. Both a Set and a Get since we need to be able to edit locally.
'
Private Property Set LocalConfig(NewLocalConfig As Dictionary)
  Dim strLocalFile As String
  strLocalFile = DefaultConfigPath & Application.Path & "local_config.json"
  ClassHelpers.WriteJson strLocalFile, NewLocalConfig
End Property

Private Property Get LocalConfig() As Dictionary
  If p_LocalConfig Is Nothing Then
    Dim strLocalFile As String
    strLocalFile = DefaultConfigPath & Application.PathSeparator & "local_config.json"
    
    If Utils.IsItThere(strLocalFile) = False Then
    ' Only user-supplied local setting right now is Region so just using a MsgBox
    ' but eventually move this to a more flexible UserForm
      Dim strMessage As String
      strMessage = "Are you in the United States?" & vbNewLine & vbNewLine & _
      vbTab & "US users: click YES" & vbNewLine & _
      vbTab & "UK users: click NO"
      
      Dim strTitle As String
      strTitle = "Please Select Your Region"
 
      Dim strRegion As String
      If MsgBox(strMessage, vbYesNo + vbQuestion, strTitle) = vbYes Then
        strRegion = "us"
      Else
        strRegion = "uk"
      End If

      Set p_LocalConfig = New Dictionary
      p_LocalConfig.Add "region", strRegion
      ClassHelpers.WriteJson strLocalFile, p_LocalConfig
    Else
      Set p_LocalConfig = ClassHelpers.ReadJson(strLocalFile)
    End If
  End If
  Set LocalConfig = p_LocalConfig
End Property
