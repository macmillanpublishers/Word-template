VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Paths_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'
' NOTES
' * Do not include path separator in any directory properties; if you need one
'   later, include it in the calling method or procedure.
' =============================================================================

' *****************************************************************************
'       CLASS VARIABLES
' *****************************************************************************

Private p_FullJSON As Dictionary
Private p_PathsJSON As Dictionary


' *****************************************************************************
'       PUBLIC PROPERTIES
' *****************************************************************************

' ===== FullJSON ==============================================================
' Dictionary object of all the data in the JSON config file. We want to only
' read the json file once, so if we need to delete the config (e.g., to replace
' it with a newer version), this instance can still run. The MacFile_ and
' MacFileDictionary_ objects will get the data from here. Note all of the
' objects in this dictionary are also dictionaries.

Public Property Get FullJSON() As Dictionary
    If p_FullJSON Is Nothing Then
        Dim strPath As String
        strPath = FullConfigPath
        Dim strJSON As String
        strJSON = SharedMacros.ReadTextFile(Path:=strPath, _
            FirstLineOnly:=False)
        Set p_FullJSON = JsonConverter.ParseJson(json_String:=strJSON)
    End If
    
    Set FullJSON = p_FullJSON
    
End Property

'' ===== LocalPath =============================================================
' Returns string of the directory a specific file group should be saved in. If
' it doesn't exist, it is created.

' At least one file MUST be in the Word STARTUP directory, which is outside of
' our project structure, so it tests for that as well.

Public Property Get LocalPath(GroupName As String) As String
    Dim strLocalPath As String

    ' If this next line works, don't need "startupFiles" in "paths" object
    If GroupName = "startupFiles" Then
        strLocalPath = Application.StartupPath
    Else
        If PathsJSON.Exists(GroupName) Then
            strLocalPath = HomeDir & Application.PathSeparator & _
            PathsJSON(GroupName)
        Else
            MsgBox "That is not a valid GroupName!"
        End If
    End If

    CheckDir (strLocalPath)
    LocalPath = strLocalPath
End Property

' ===== LogsDir ===============================================================
' Path to logs will always be the same for each file, so we won't get it from
' MacFile_ directly.

Public Property Get LogsDir() As String
    LogsDir = LocalPath(GroupName:="logFiles")
End Property

' ===== TempPath ==========================================================
' Path to OS-specific local temp directory.

Public Property Get TempPath() As String
    Dim strTmpDir As String
    #If Mac Then
        strTmpDir = MacScript("path to temporary items as string")
    #Else
        strTmpDir = Environ("TEMP")
    #End If

    ' Not 100% confident these return trailing separator consistently, so:
    If Right(strTmpDir, 1) = Application.PathSeparator Then
        strTmpDir = Left(strTmpDir, Len(strTmpDir) - 1)
    End If
    
    ' Don't use CheckDir here because shouldn't create new Tmp dir
    If IsItThere(strTmpDir) = True Then
        TempPath = strTmpDir
    Else
        MsgBox "Um, your Temp folder is missing?"
    End If
End Property

' ===== DownloadURL ===========================================================
' Full URL to download the file from, excluding the file name and trailing
' separator. Will be assigned to WebClient.BaseURL.

Public Property Get DownloadURL() As String
    Dim dictDownloads As Dictionary
    Set dictDownloads = FullJSON("downloads")

    DownloadURL = dictDownloads("baseURL") & "/" & _
        dictDownloads(dictDownloads("current")) & "/"
End Property


' *****************************************************************************
'       PRIVATE PROPERTIES
' *****************************************************************************

' ===== FullConfigPath ========================================================
' Takes full local path to config file as input (incl. name and extension).
' Writes it to CustomDocumentProperty for document to read later. All other
' objects should get this from here, not independently if possible.

Private Property Let FullConfigPath(strFullConfigPath As String)
    ThisDocument.CustomDocumentProperties("FullConfigPath").value = _
        strFullConfigPath
End Property

Private Property Get FullConfigPath() As String
    FullConfigPath = ThisDocument.CustomDocumentProperties("FullConfigPath")
End Property

' ===== PathsJSON =============================================================
' Dictionary object of the "paths" object in the config JSON. Contains names of
' all the assorted directories we need. Doesn't do anything with them; but want
' to pull this dictionary from FullJSON in a single place in case we need to
' change something later.

Private Property Get PathsJSON() As Dictionary
    If p_PathsJSON Is Nothing Then
        Set p_PathsJSON = FullJSON("paths")
    End If
    PathsJSON = p_PathsJSON
End Property

' ===== RootDir ===============================================================
' Returns string representing the root directory of the project files (not the
' system root). Project should be saved to some level of the user directory
' because security at other levels can vary for macro files. Right now the only
' option is "user", but I've set it up as a Select statement so it will be
' relatively easy to add other options later. If the directory doesn't exist,
' it throws an error. Should we create a fallback?

Private Property Get RootDir() As String
    Dim strRootDir As String

    Select Case PathsJSON("root")
        Case "macmillan"
            #If Mac Then
                ' This returns a string with the trailing path separator
                ' included so we need to remove it.

                ' Also Matt suggested Users/username/Library/Application Support
                ' as a place less likely to be accidentally deleted.
                strRootDir = MacScript("return (path to home folder) as string")
                If Right(strRootDir, 1) = Application.PathSeparator Then
                    strRootDir = Left(strRootDir, Len(strRootDir) - 1)
                End If
            #Else
                strRootDir = Environ("HOMEDRIVE") & Environ("HOMEPATH")
            #End If
        Case Else
            MsgBox "That's not a valid value for Root directory."
    End Select
    
    ' Don't use CheckDir, don't try to create home directory...
    If IsItThere(strRootDir) = False Then
        MsgBox "You don' seem to have a home directory..."
    End If
    
    RootDir = strRootDir

End Property

' ===== HomeDir ===============================================================
' Returns string of full path to the project home directory, i.e. where all of
' the Word template files are saved. If it doesn't exist, it will create it, so
' calling procedure doesn't need to test for existence.

Private Property Get HomeDir() As String
    Dim strHomeDir As String
    strHomeDir = RootDir & Application.PathSeparator & PathsJSON("home")
    CheckDir (strHomeDir)
    HomeDir = strHomeDir
End Property


' *****************************************************************************
'       PRIVATE METHODS
' *****************************************************************************

' ===== Check =================================================================
' Parameter is string representing full path to a directory; checks if the dir
' exists and if not, creates it. Throws and error if it can't create (e.g.,
' root project dir). Making it its own method (with a pretty generic name) so
' it's easy to add additional checks later if we need to.

Private Sub CheckDir(MyDir As String)
    On Error GoTo CheckDirError
    If IsItThere(MyDir) = False Then
        MkDir MyDir
    End If
    On Error GoTo 0
    Exit Sub
CheckDirError:
    MsgBox "Nope! Looks like " & MyDir & " doesn't exist and we can't create it" _
        & vbNewLine & "Error " & Err.Number & ": " & Err.Description
End Sub


' *****************************************************************************
'       CLASS EVENTS
' *****************************************************************************

' ===== Initialize ============================================================
' This is a static class, so Initialize will be called the first time the
' default object is called.

Private Sub Class_Initialize()

End Sub
