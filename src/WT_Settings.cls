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
Private p_GlobalConfig As Dictionary
Private p_DefaultConfigPath As String


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
' Are we in debugging mode?

Public Property Get DebugOn() As Boolean
  If Overrides.Exists("DebugOn") Then
    DebugOn = True
  Else
    DebugOn = False
  End If
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

  If Overrides("DownloadBranch").Exists(Repo) Then
    DownloadBranch = Overrides("DownloadBranch")(Repo)
  Else
    DownloadBranch = "master"
  End If
End Property

' ===== FileInfo ==============================================================
' Processes data in various config files to return info about a specific file
' required by the template.

' RETURNS
' Dictionary with the following keys:
' "Final" : Full path (with file name) to file on local computer
' "Tmp" : Full path (with file name) to file in TMP dir
' "Log" : Full path to download log file for this file
' "Remote" : Full download URL for file from remote repo/release

Public Property Get FileInfo(FileName As String) As Dictionary
  Dim dictWorkingData As Dictionary
  Set dictWorkingData = GetWorkingData(FileName)

  Dim strFinal As String
  Dim strTmp As String
  Dim strLog As String
  Dim strRemote As String
  Dim strBaseName As String
  
  strBaseName = Utils.GetFileNameOnly(FileName)

  strFinal = BuildPath(FileName)
  strTmp = TmpDir & Application.PathSeparator & FileName
  strLog = RootDir & Application.PathSeparator & "log" & Application.PathSeparator _
    & strBaseName & "_updates.log"
  strRemote = BuildPath(FileName, True)

  Set FileInfo = New Dictionary
  With FileInfo
    .Add "Final", strFinal
    .Add "Tmp", strTmp
    .Add "Log", strLog
    .Add "Remote", strRemote
  End With
  
End Property


' ***** PRIVATE PROPERTIES *****************************************************

' ===== RootDir ===============================================================
' Local root directory for all this stuff. Is

' ===== TmpDir ================================================================
' Returns path to TEMP directory, no trailing path separator

Private Property Get TmpDir() As String
  Dim strTmpDir As String
  #If Mac Then
    strTmpDir = MacScript("path to temporary items as string")
  #Else
    strTmpDir = Environ("TEMP")
  #End If
' Remove trailing path separator, if any
  If Right(strTmpDir, 1) = Application.PathSeparator Then
    strTmpDir = Left(strTmpDir, Len(strTmpDir) - 1)
  End If
  TmpDir = strTmpDir
End Property


' ===== DefaultConfigPath =====================================================
' Default path to config files on local machine. Since file paths are stored in
' the config files, we need a way to access the config files themselves before
' we can read them.

Private Property Get DefaultConfigPath() As String
  If p_DefaultConfigPath = vbNullString Then
    p_DefaultConfigPath = ThisDocument.Path & Application.PathSeparator & "config"
    If Utils.IsItThere(p_DefaultConfigPath) = False Then
      MkDir p_DefaultConfigPath
    End If
  End If
    DefaultConfigPath = p_DefaultConfigPath
End Property

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

' ***** PRIVATE METHODS *******************************************************

' ===== GetWorkingData ========================================================
' Loop through all config files in order of least important to more important
' and add each value to working dictionary. Items that take precedence will
' overwrite the value from the previous configs.

Private Function GetWorkingData(FileName As String) As Dictionary
  Dim collConfigs As Collection
  Dim objDict As Object
  Dim varKey1 As Variant
  Dim dictFileData As Dictionary
  Dim varKey2 As Variant
  Dim dictWorkingData As Dictionary

  Set dictWorkingData = New Dictionary

  If FileName = "global_config.json" Then
  ' Read from CustomDocumentProperties, since we can't get data
  ' from a file we don't yet have
    With dictWorkingData
      .Add "source_urls", Utils.ToCollection(Split(ThisDocument.CustomDocumentProperties("source_urls"), ","))
      .Add "source", ThisDocument.CustomDocumentProperties("source")
      .Add "organization", ThisDocument.CustomDocumentProperties("organization")
      .Add "repo", ThisDocument.CustomDocumentProperties("repo")
      .Add "subfolders", Utils.ToCollection(Split(ThisDocument.CustomDocumentProperties("subfolders"), ","))
    End With
  
  Else
  ' Read data from config files

    Set collConfigs = New Collection
    collConfigs.Add GlobalConfig
    collConfigs.Add RegionConfig
    collConfigs.Add LocalConfig
  
    For Each objDict In collConfigs
      For Each varKey1 In objDict.Keys
        If varKey1 = "files" Then
          Set dictFileData = objDict("files")(FileName)
          For Each varKey2 In dictFileData.Keys
            If IsObject(dictFileData(varKey2)) = True Then
              Set dictWorkingData.Item(varKey2) = dictFileData(varKey2)
            Else
              dictWorkingData.Item(varKey2) = dictFileData(varKey2)
            End If
          Next varKey2
        Else
          If IsObject(objDict(varKey1)) = True Then
            Set dictWorkingData.Item(varKey1) = objDict(varKey1)
          Else
            dictWorkingData.Item(varKey1) = objDict(varKey1)
          End If
        End If
      Next varKey1
    Next objDict
  End If

  Set GetWorkingData = dictWorkingData
End Function

' ===== BuildPath =============================================================
' The part of the path to the file after the root directory/URL, so we can match
' the local path to the remote URL path.

Private Function BuildPath(FileName As String, Optional IsUrl As Boolean _
  = False) As String
  Dim dictWorkingData As Dictionary
  Set dictWorkingData = GetWorkingData(FileName)

' Add path elements to a collection, we'll Join later
  Dim collPath As Collection
  Set collPath = New Collection

' Set separator and base for local files vs. URLs
  Dim strSeparator As String
  If IsUrl = True Then
    strSeparator = "/"
    collPath.Add dictWorkingData("source_urls")(dictWorkingData("source"))
  Else
    strSeparator = Application.PathSeparator
    collPath.Add RootDir
  End If

' Add next two elements (all same format)
  collPath.Add dictWorkingData("organization")
  collPath.Add dictWorkingData("repo")

' Add middle parts based on type of path we're building
  If IsUrl = True Then
    If dictWorkingData("source") = "releases" Then
      collPath.Add "releases"
      collPath.Add "download"
      collPath.Add dictWorkingData("current_release")
    Else
      If dictWorkingData.Exists("branch") = True Then
        collPath.Add dictWorkingData("branch")
      Else
        collPath.Add "master"
      End If
    End If
  End If

' Add subfolders if we have any (and not downloading from releases)
  If dictWorkingData.Exists("subfolders") = True Then
    If IsUrl = False Or dictWorkingData("source") = "direct" Then
      Dim collSubfolders As Collection
      Set collSubfolders = dictWorkingData("subfolders")

      Dim varDir As Variant
      For Each varDir In collSubfolders
        collPath.Add varDir
      Next varDir
    End If
  End If

' Add the file name to the end!
  collPath.Add FileName
  
' No native Join function for Collections, so convert to an array first
  Dim varPathArray As Variant
  varPathArray = Utils.ToArray(collPath)
  
  BuildPath = Join(varPathArray, strSeparator)
End Function
