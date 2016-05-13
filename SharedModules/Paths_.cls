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
'       CLASS CONSTANTS
' *****************************************************************************
' For error handling
Private Const strPaths As String = "Paths_."
Private Const strConfigKey As String = "FullConfigPath"


' *****************************************************************************
'       CLASS VARIABLES
' *****************************************************************************

Private p_FullJSON As Dictionary
Private p_PathsJSON As Dictionary
Private p_DownloadURL As String
Private p_FullConfigPath As String
Private p_HomeDir As String
Private p_LocalPath As String
Private p_LogsDir As String
Private p_RootDir As String
Private p_TempPath As String


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
    On Error GoTo FullJSONGetError
    If p_FullJSON Is Nothing Then
        Dim strPath As String
        strPath = FullConfigPath
        Dim strJSON As String
        strJSON = SharedMacros_.ReadTextFile(Path:=strPath, _
            FirstLineOnly:=False)
        Set p_FullJSON = JsonConverter.ParseJson(json_String:=strJSON)
    End If
    
    Set FullJSON = p_FullJSON
FullJSONGetFinish:
    Exit Property

FullJSONGetError:
    Err.Source = Err.Source & strPaths & "FullJSON.Get"
    If ErrorChecker(Err) = False Then
        Resume
    Else
        Resume FullJSONGetFinish
    End If
End Property

'' ===== LocalPath =============================================================
' Returns string of the directory a specific file group should be saved in. If
' it doesn't exist, it is created.

' At least one file MUST be in the Word STARTUP directory, which is outside of
' our project structure, so it tests for that as well.

Public Property Get LocalPath(GroupName As String) As String
    On Error GoTo LocalPathGetError
    If p_LocalPath = vbNullString Then
        ' If this next line works, don't need "startupFiles" in "paths" object
        If GroupName = "startupFiles" Then
            p_LocalPath = Application.StartupPath
        Else
            If PathsJSON.Exists(GroupName) Then
                p_LocalPath = HomeDir & Application.PathSeparator & _
                PathsJSON(GroupName)
            Else
                Err.Raise MacError.err_GroupNameInvalid
            End If
        End If
    End If

    CheckDir (p_LocalPath)
    LocalPath = p_LocalPath
LocalPathGetFinish:
    Exit Property

LocalPathGetError:
    Err.Source = Err.Source & strPaths & "LocalPath.Get"
    If ErrorChecker(Err) = False Then
        Resume
    Else
        Resume LocalPathGetFinish
    End If
End Property

' ===== LogsDir ===============================================================
' Path to logs will always be the same for each file, so we won't get it from
' MacFile_ directly.

Public Property Get LogsDir() As String
    On Error GoTo LogsDirGetError
    If p_LogsDir = vbNullString Then
        ' LocalPath calls CheckDir, which will create if not extant
        p_LogsDir = LocalPath(GroupName:="logFiles")
    End If
    LogsDir = p_LogsDir
LogsDirGetFinish:
    Exit Property

LogsDirGetError:
    Err.Source = Err.Source & strPaths & "LogsDir.Get"
    If ErrorChecker(Err) = False Then
        Resume
    Else
        Resume LogsDirGetFinish
    End If
End Property

' ===== TempPath ==========================================================
' Path to OS-specific local temp directory.

Public Property Get TempPath() As String
    On Error GoTo TempPathError
    If p_TempPath = vbNullString Then
        #If Mac Then
            p_TempPath = MacScript("path to temporary items as string")
        #Else
            p_TempPath = Environ("TEMP")
        #End If
    
        ' Not 100% confident these return trailing separator consistently, so:
        If Right(p_TempPath, 1) = Application.PathSeparator Then
            p_TempPath = Left(p_TempPath, Len(p_TempPath) - 1)
        End If
    End If

    ' Don't use CheckDir here because shouldn't create new Tmp dir
    If IsItThere(p_TempPath) = True Then
        TempPath = p_TempPath
    Else
        Err.Raise MacError.err_TempMissing
    End If
TempPathFinish:
    Exit Property

TempPathError:
    Err.Source = Err.Source & strPaths & "TempPath.Get"
    If ErrorChecker(Err, p_TempPath) = False Then
        Resume
    Else
        Resume TempPathFinish
    End If
End Property

' ===== DownloadURL ===========================================================
' Full URL to download the file from, excluding the file name and trailing
' separator. Will be assigned to WebClient.BaseURL.

Public Property Get DownloadURL() As String
    On Error GoTo DownloadURLError
    If p_DownloadURL = vbNullString Then
        Dim dictDownloads As Dictionary
        Set dictDownloads = FullJSON("downloads")
    
        p_DownloadURL = dictDownloads("baseURL") & "/" & _
            dictDownloads(dictDownloads("current"))
    End If
    DownloadURL = p_DownloadURL
DownloadURLFinish:
    Exit Property

DownloadURLError:
    Err.Source = Err.Source & strPaths & "DownloadURL.Get"
    If ErrorChecker(Err) = False Then
        Resume
    Else
        Resume DownloadURLFinish
    End If
End Property


' *****************************************************************************
'       PRIVATE PROPERTIES
' *****************************************************************************

' ===== FullConfigPath ========================================================
' Takes full local path to config file as input (incl. name and extension).
' Writes it to CustomDocumentProperty for document to read later. All other
' objects should get this from here, not independently, if possible. DON'T
' check if it's there, because we need this path in order to save the file.

Private Property Let FullConfigPath(strFullConfigPath As String)
    On Error GoTo FullConfigPathLetError
    p_FullConfigPath = strFullConfigPath
    ' Also update the actual doc properties

    If SharedMacros_.DocPropExists(Me, strConfigKey) = True Then
        ThisDocument.CustomDocumentProperties(strConfigKey).Value = _
            p_FullConfigPath
    Else
        ThisDocument.CustomDocumentProperties.Add _
            Name:=strConfigKey, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeString, _
            Value:=p_FullConfigPath
    End If
FullConfigPathLetFinish:
    Exit Property

FullConfigPathLetError:
    Err.Source = Err.Source & strPaths & "FullConfigPath.Get"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume FullConfigPathLetFinish
    End If
End Property

Private Property Get FullConfigPath() As String
    On Error GoTo FullConfigPathGetError
    If p_FullConfigPath = vbNullString Then
        If SharedMacros_.DocPropExists(Me, strConfigKey) = True Then
            p_FullConfigPath = ThisDocument.CustomDocumentProperties(strConfigKey)
        Else
            Err.Raise MacError.err_ConfigPathNull
        End If
    End If
    FullConfigPath = p_FullConfigPath
FullConfigPathGetFinish:
    Exit Property

FullConfigPathGetError:
    Err.Source = Err.Source & strPaths & "FullConfigPath.Get"
    Debug.Print Err.Number & ": " & Err.Description
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume FullConfigPathGetFinish
    End If
End Property

' ===== PathsJSON =============================================================
' Dictionary object of the "paths" object in the config JSON. Contains names of
' all the assorted directories we need. Doesn't do anything with them; but want
' to pull this dictionary from FullJSON in a single place in case we need to
' change something later.

Private Property Get PathsJSON() As Dictionary
    On Error GoTo PathsJsonGetError
    If p_PathsJSON Is Nothing Then
        Set p_PathsJSON = FullJSON("paths")
    End If
    Set PathsJSON = p_PathsJSON
PathsJsonGetFinish:
    Exit Property

PathsJsonGetError:
    Err.Source = Err.Source & strPaths & "PathsJSON.Get"
    If ErrorChecker(Err) = False Then
        Resume
    Else
        Resume PathsJsonGetError
    End If
End Property

' ===== RootDir ===============================================================
' Returns string representing the root directory of the project files (not the
' system root). Project should be saved to some level of the user directory
' because security at other levels can vary for macro files. Right now the only
' option is "user", but I've set it up as a Select statement so it will be
' relatively easy to add other options later. If the directory doesn't exist,
' it throws an error. Should we create a fallback?

Private Property Get RootDir() As String
    On Error GoTo RootDirGetError
    Dim strRootCode As String
    If p_RootDir = vbNullString Then
        strRootCode = PathsJSON("root")
        Select Case strRootCode
            Case "macmillan"
                #If Mac Then
                    ' This returns a string with the trailing path separator
                    ' included so we need to remove it.
    
                    ' Also Matt suggested Users/username/Library/Application Support
                    ' as a place less likely to be accidentally deleted.
                    ' p_RootDir = MacScript("return (path to application" & _
                    " support) as string")
                    
                    p_RootDir = MacScript("return (path to library folder " & _
                        "from user domain) as string")
                #Else
                    ' which is actually C:\Users\%USERNAME%\AppData\Roaming
                    ' as recommended by Rakesh
                    p_RootDir = Environ("APPDATA")
                #End If
            Case Else
                Err.Raise MacError.err_RootDirInvalid
        End Select
        ' Make sure no trailing path separator
        If Right(p_RootDir, 1) = Application.PathSeparator Then
            p_RootDir = Left(p_RootDir, Len(p_RootDir) - 1)
        End If
    End If
    
    ' Don't use CheckDir, don't try to create root directory...
    If IsItThere(p_RootDir) = False Then
        Err.Raise MacError.err_RootDirMissing
    End If
    
    RootDir = p_RootDir
RootDirGetFinish:
    Exit Property

RootDirGetError:
    Err.Source = Err.Source & strPaths & "RootDir.Get"
    Dim strErrorValue As String

    Select Case Err.Number
        Case MacError.err_RootDirInvalid
            strErrorValue = strRootCode
        Case MacError.err_RootDirMissing
            strErrorValue = p_RootDir
        Case Else
            strErrorValue = vbNullString
    End Select

    If SharedMacros_.ErrorChecker(Err, strErrorValue) = False Then
        Resume
    Else
        Resume RootDirGetFinish
    End If
End Property

' ===== HomeDir ===============================================================
' Returns string of full path to the project home directory, i.e. where all of
' the Word template files are saved. If it doesn't exist, CheckDir will create
' it, so calling procedure doesn't need to test for existence.

Private Property Get HomeDir() As String
    On Error GoTo HomeDirGetError
    If p_HomeDir = vbNullString Then
        p_HomeDir = RootDir & Application.PathSeparator & PathsJSON("home")
    End If
    CheckDir p_HomeDir
    HomeDir = p_HomeDir
HomeDirGetFinish:
    Exit Property

HomeDirGetError:
    Err.Source = Err.Source & strPaths & "HomeDir.Get"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume HomeDirGetFinish
    End If
End Property

' *****************************************************************************
'       PUBLIC METHODS
' *****************************************************************************

' ===== CleanUp ===============================================================
' Static class will never call a Class_Terminate event, so we need to create
' our own and remember to call it when the original sub is complete. Hopefully
' can find a more reliable way to do this in the future. But we need to
' explicitly set the object properties to Nothing or they will persist, even
' beyond the end of the macro sometimes.

' And for fun let's reset all properties; might need to call before calling
' procedures completes if we download a new config file. Can break out into
' separate Property Let procedures later if we need finer control.

Public Sub ClearPaths()
    On Error GoTo ClearPathsError
    If Not p_FullJSON Is Nothing Then
        Set p_FullJSON = Nothing
    End If
    
    If Not p_PathsJSON Is Nothing Then
        Set p_PathsJSON = Nothing
    End If
    
    p_DownloadURL = vbNullString
    p_FullConfigPath = vbNullString
    p_HomeDir = vbNullString
    p_LocalPath = vbNullString
    p_LogsDir = vbNullString
    p_RootDir = vbNullString
    p_TempPath = vbNullString
ClearPathsFinish:
    Exit Sub

ClearPathsError:
    Err.Source = Err.Source & strPaths & "ClearPaths"
    If SharedMacros_.ErrorChecker(Err) = False Then
        Resume
    Else
        Resume ClearPathsFinish
    End If
End Sub


' *****************************************************************************
'       PRIVATE METHODS
' *****************************************************************************

' ===== Check =================================================================
' Parameter is string representing full path to a directory; checks if the dir
' exists and if not, creates it. Throws and error if it can't create (e.g.,
' root project dir). Making it its own method (with a pretty generic name) so
' it's easy to add additional checks later if we need to. Returns True if dir
' exists.

Private Function CheckDir(MyDir As String) As Boolean
    On Error GoTo CheckDirError
    If SharedMacros_.IsItThere(MyDir) = False Then
        ' Mac Word 2011 can't handle file path > 32 char
        If SharedMacros_.IsOldMac = True Then
            Dim strCommand As String
            strCommand = MacScript("return quoted form of posix path of " & MyDir)
            strCommand = "mkdir -p " & strCommand
            SharedMacros_.ShellAndWaitMac (strCommand)
        Else
            MkDir MyDir
        End If
        ' Make sure it worked
        CheckDir = SharedMacros_.IsItThere(MyDir)
    Else
        CheckDir = True
    End If
CheckDirFinish:
    Exit Function

CheckDirError:
    Err.Source = Err.Source & strPaths & "CheckDir " & MyDir
    If SharedMacros_.ErrorChecker(Err, MyDir) = False Then
        Resume
    Else
        Resume CheckDirFinish
    End If
End Function


' *****************************************************************************
'       CLASS EVENTS
' *****************************************************************************

' ===== Terminate =============================================================
' Psych! You can't call Class_Terminate on a static class.
