Attribute VB_Name = "FileInstaller"
Option Explicit
Option Base 1

Sub Installer(Installer As Boolean, TemplateName As String, ByRef FileName() As String, ByRef FinalDir() As String)
'created by Erica Warren - erica.warren@macmillan.com

'======== PURPOSE ===================================
'Downloads and installs an array of template files & logs the downloads

'======== USE =======================================
'This is Part 2 of 2. Must be called from a sub in another module that declares file names and locations.
'The template file needs to be uploaded as an attachment to https://confluence.macmillan.com/display/PBL/Test
'If this is an installer file, The Part 1 code needs to reside in the ThisDocument module as a sub called
'Documents_Open in a .docm file so that it will launch when users open the file.

'"Installer" argument = True if this is for a standalone installtion file.
'"Installer" argument = False is this is part of a daily check of the current file and only updates if out of date.
    
    '' --------------- Check that variables were passed correctly -------------------------
    'Dim x As Long
    'For x = LBound(FileName()) To UBound(FileName())
    '    Debug.Print & " " & FileName(x) & " " & FinalDir(x) & vbNewLine
    'Next x
    
    '' --------------- Set up variable names ----------------------------------------------
    '' Create style directory and logfile names
    Dim a As Long
    Dim arrLogInfo() As Variant
    ReDim arrLogInfo(1 To 3)
    Dim strStyleDir() As String
    ReDim strStyleDir(LBound(FileName()) To UBound(FileName()))
    Dim strLogDir() As String
    ReDim strLogDir(LBound(FileName()) To UBound(FileName()))
    Dim strFullLogPath As String
    ReDim strFullLogPath(LBound(FileName()) To UBound(FileName()))

    ' ------------ Define Log Dirs and such -----------------------------------------
    For a = LBound(FileName()) To UBound(FileName())
        arrLogInfo() = CreateLogFileInfo(FileName)
        strStyleDir(a) = arrLogInfo(1)
        strLogDir(a) = arrLogInfo(2)
        strFullLogPath(a) = arrLogInfo(3)
    Next a
    
    ' ------------- Check if we need to do an installation ---------------------------
    ' Check if template exists
    Dim installCheck() As Boolean
    ReDim installCheck(LBound(FileName()) To UBound(FileName()))
    Dim blnTemplateExists() As Boolean
    ReDim blnTemplateExists(LBound(FileName()) To UBound(FileName()))
    Dim blnLogUpToDate() As Boolean
    ReDim blnLogUpToDate(LBound(FileName()) To UBound(FileName()))
    Dim b As Long
    
    For b = LBound(FileName()) To UBound(FileName())
    
        'Check if log dir/file exists, create if it doesn't, check last mod date if it does
        ' If last mod date less than 1 day ago, CheckLog = True
        blnLogUpToDate(b) = CheckLog(strStyleDir, strLogDir, strFullLogPath(b))
        'Debug.Print FileName(b) & " log exists and was checked today: " & blnLogUpToDate(b)
        
        ' Check if template exists, if not create any missing directories
        blnTemplateExists(b) = IsTemplateThere(FinalDir(b), FileName(b), strFullLogPath(b))
        'Debug.Print FileName(b) & " exists: " & blnTemplateExists(b)
                
        If Installer = False Then 'Because if it's an installer, we just want to install the file
            If blnLogUpToDate(b) = True And blnTemplateExists(b) = True Then ' already checked today, already exists
                installCheck(b) = False
            ElseIf blnLogUpToDate(b) = False And blnTemplateExists(b) = True Then 'Log is new or not checked today, already exists
                'check version number
                installCheck(b) = NeedUpdate(FinalDir(b), FileName(b), strFullLogPath(b))
            Else ' blnTemplateExists = False, just download new template
                 installCheck(b) = True
            End If
        Else
            installCheck(b) = True
        End If
        
    Next b

    ' ---------------- Create new array of template files we need to install -----------------
    Dim strInstallFile() As String
    Dim strInstallDir() As String
    Dim c As Long
    Dim x As Long
    
    x = 0
    
    For c = LBound(FileName()) To UBound(FileName())
        If installCheck(c) = True Then
            x = x + 1
            ReDim Preserve strInstallFile(1 To x)
                strInstallFile(x) = FileName(c)
            ReDim Preserve strInstallDir(1 To x)
                strInstallDir(x) = FinalDir(c)
        End If
    Next c
    
    'Debug.Print strInstallFile(1) & vbNewLine & strInstallDir(1)
    
    ' ---------------- Check if new array is allocated -----------------------------------
    If IsArrayEmpty(strInstallFile()) = True Then
        If Installer = True Then
            Application.Quit (wdDoNotSaveChanges)
        Else
            Exit Sub
        End If
    Else ' There are values in the array and we need to install them
    
        ' Alert user that installation is happening
        Dim strWelcome As String
        If Installer = True Then
            strWelcome = "Welcome to the " & TemplateName & " Installer!" & vbNewLine & vbNewLine & _
                "Please click OK to begin the installation. It should only take a few seconds."
        Else
            strWelcome = "Your " & TemplateName & " is out of date. Click OK to update automatically."
        End If
    
        If MsgBox(strWelcome, vbOKCancel, TemplateName) = vbCancel Then
            MsgBox "Please try to install the files at a later time."
            
            If Installer = True Then
                Application.Quit (wdDoNotSaveChanges)
            End If
            
            Exit Sub
        End If
    End If
    
    ' ---------------- Close any open docs (with prompt) -----------------------------------
    Call CloseOpenDocs
        
    '----------------- download template files ------------------------------------------
    Dim logString As String
    Dim d As Long
    
    For d = LBound(strInstallFile()) To UBound(strInstallFile())
        'If False, error in download; user was notified in DownloadFromConfluence function
        If DownloadFromConfluence(strInstallDir(d), strFullLogPath(d), strInstallFile(d)) = False Then
            If Installer = True Then
                Application.Quit (wdDoNotSaveChanges)
            Else
                Exit Sub
            End If
        End If
    Next d
    
    '------Display installation complete message and close doc (ending sub)---------------
    Dim strComplete As String
    
    strComplete = "The " & TemplateName & " has been installed on your computer." & vbNewLine & vbNewLine & _
        "When you restart Word, the template will be available."
        
    MsgBox strComplete, vbOKOnly, "Installation Successful"
    
    '------Close and restart Word for template changes to take effect---------------------
    'Would love to get restart to work...
    'Dim restartTime As Variant
    'restartTime = Now + TimeValue("00:00:01")
    'Application.OnTime When:=restartTime, Name:="Restart"
    If Installer = True Then
        Application.Quit SaveChanges:=wdDoNotSaveChanges          'DEBUG: comment out this line
    End If
        
End Sub

Private Function IsTemplateThere(Directory As String, FileName As String, Log As String)
    
    'Create full path to template
    Dim strFullPath As String
    Dim logString As String
    strFullPath = Directory & Application.PathSeparator & FileName
    
    '------------------------- Check if template exists ----------------------------------
    ' Create directory if it doesn't exist
    If IsItThere(Directory) = False Then
        MkDir (Directory)
        IsTemplateThere = False
        logString = Now & " -- Creating template directory."
    Else
        ' Check if template file exists
        If IsItThere(strFullPath) = False Then
            IsTemplateThere = False
            logString = Now & " -- " & FileName & " doesn't exist in " & Directory
        Else
            IsTemplateThere = True
            logString = Now & " -- " & FileName & " already exists."
        End If
    End If

    LogInformation Log, logString
End Function

Private Function NeedUpdate(Directory As String, FileName As String, Log As String) As Boolean
'Directory argument should be the final directory the template should go in.
'File should be the template file name
'Log argument should be full path to log file

    '------------------------- Get installed version number -----------------------------------
    Dim logString As String
    Dim strFullTemplatePath As String
    
    strFullTemplatePath = Directory & Application.PathSeparator & FileName
    'Debug.Print "NeedUpdate Path: " & strFullTemplatePath
    
    'Get version number of installed template
    Dim strInstalledVersion As String
    If IsItThere(strFullTemplatePath) = True Then
        Documents.Open FileName:=strFullTemplatePath, ReadOnly:=True ', Visible:=False      'Mac Word 2011 doesn't allow Visible as an argument :(
        strInstalledVersion = Documents(strFullTemplatePath).CustomDocumentProperties("Version")
        Documents(strFullTemplatePath).Close
        logString = Now & " -- installed version is " & strInstalledVersion
    Else
        strInstalledVersion = 0     ' Template is not installed
        logString = Now & " -- No template installed, version number is 0."
    End If
    
    LogInformation Log, logString
    
    '------------------------- Try to get current version's number from Confluence ------------
    Dim strVersion As String
    Dim strFullVersionPath As String
    
    'Debug.Print FileName
    'Debug.Print InStrRev(FileName, ".do")
    strVersion = Left(FileName, InStrRev(FileName, ".do") - 1)
    strVersion = strVersion & ".txt"
    strFullVersionPath = Directory & Application.PathSeparator & strVersion
    'Debug.Print strVersion
    
    'If False, error in download; user was notified in DownloadFromConfluence function
    If DownloadFromConfluence(Directory, Log, strVersion) = False Then
        NeedUpdate = False
    End If
        
    '-------------------- Get version number of current template ---------------------
    If IsItThere(strFullVersionPath) = True Then
        NeedUpdate = True
        Dim strCurrentVersion As String
        strCurrentVersion = ImportVariable(strFullTemplatePath)
        logString = Now & " -- Current version is " & strCurrentVersion
    Else
        NeedUpdate = False
        logString = Now & " -- Download of version file for " & FileName & " failed."
    End If
        
    LogInformation Log, logString
    
    If NeedUpdate = False Then
        Exit Function
    End If
    
    '--------------------- Compare version numbers -----------------------------------
    
    If strInstalledVersion >= strCurrentVersion Then
        NeedUpdate = False
        logString = Now & " -- Current version matches installed version."
    Else
        NeedUpdate = True
        logString = Now & " -- Current version greater than installed version."
    End If
    
    LogInformation Log, logString
    
End Function

Private Sub CloseOpenDocs()

    '-------------Check for/close open documents---------------------------------------------
    Dim strInstallerName As String
    Dim strSaveWarning As String
    Dim objDocument As Document
    Dim b As Long
    
    strInstallerName = ThisDocument.Name
        'Debug.Print "Installer Name: " & strInstallerName
        'Debug.Print "Open docs: " & Documents.Count
        
    If Documents.Count > 1 Then
        strSaveWarning = "All other Word documents must be closed to run the installer." & vbNewLine & vbNewLine & _
            "Click OK and I will save and close your documents." & vbNewLine & _
            "Click Cancel to exit without installing and close the documents yourself."
        If MsgBox(strSaveWarning, vbOKCancel, "Close documents?") = vbCancel Then
            ActiveDocument.Close
            Exit Sub
        Else
            For b = 1 To Documents.Count
                'Debug.Print "Current doc " & b & ": " & Documents(b).Name
                On Error Resume Next        'To skip error if user is prompted to save new doc and clicks Cancel
                    If Documents(b).Name <> strInstallerName Then       'But don't close THIS document
                        Documents(b).Save   'separate step to trigger Save As prompt for previously unsaved docs
                        Documents(b).Close
                    End If
                On Error GoTo 0
            Next b
        End If
    End If
    
End Sub

Private Function ImportVariable(strFile As String) As String
 
    Open strFile For Input As #1
    Line Input #1, ImportVariable
    Close #1
 
End Function

Private Function IsArrayEmpty(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' By Chip Pearson, http://www.cpearson.com/excel/vbaarrays.htm
'
' IsArrayEmpty
' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is really the reverse of IsArrayAllocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim LB As Long
    Dim UB As Long
    
    Err.Clear
    On Error Resume Next
    If IsArray(Arr) = False Then
        ' we weren't passed an array, return True
        IsArrayEmpty = True
    End If
    
    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    UB = UBound(Arr, 1)
    If (Err.Number <> 0) Then
        IsArrayEmpty = True
    Else
        ''''''''''''''''''''''''''''''''''''''''''
        ' On rare occassion, under circumstances I
        ' cannot reliably replictate, Err.Number
        ' will be 0 for an unallocated, empty array.
        ' On these occassions, LBound is 0 and
        ' UBoung is -1.
        ' To accomodate the weird behavior, test to
        ' see if LB > UB. If so, the array is not
        ' allocated.
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        LB = LBound(Arr)
        If LB > UB Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If

End Function
