Attribute VB_Name = "SharedFileInstaller"
' created by Erica Warren - erica.warren@macmillan.com

' ======== PURPOSE ===================================
' Downloads and installs an array of template files & logs the downloads

' ======== DEPENDENCIES =======================================
' This is Part 2 of 2. Must be called from a sub in another module that declares file names and locations.
' The template file needs to be uploaded as an attachment to https://confluence.macmillan.com/display/PBL/Test
' If this is an installer file, The Part 1 code needs to reside in the ThisDocument module as a sub called
' Documents_Open in a .docm file so that it will launch when users open the file.
' Requires the SharedMacros module be installed in the same template.

Option Explicit
Option Base 1

Sub Installer(Staging As Boolean, Installer As Boolean, TemplateName As String, ByRef FileName() As String, ByRef FinalDir() As String)

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
    Dim strFullLogPath() As String
    ReDim strFullLogPath(LBound(FileName()) To UBound(FileName()))
    
    ' ------------ Define Log Dirs and such -----------------------------------------
    For a = LBound(FileName()) To UBound(FileName())
        arrLogInfo() = CreateLogFileInfo(FileName(a))
        strStyleDir(a) = arrLogInfo(1)
        strLogDir(a) = arrLogInfo(2)
        strFullLogPath(a) = arrLogInfo(3)
    Next a
    
    'Debug.Print "Style Dir is: " & strStyleDir(1) & vbNewLine & _
                "Log dir is: " & strLogDir(1) & vbNewLine & _
                "Full path to log is: " & strFullLogPath(1)
                
    ' ------------- Check if we need to do an installation ---------------------------
    ' Check if template exists
    Dim installCheck() As Boolean
    ReDim installCheck(LBound(FileName()) To UBound(FileName()))
    Dim blnTemplateExists() As Boolean
    ReDim blnTemplateExists(LBound(FileName()) To UBound(FileName()))
    Dim blnLogUpToDate() As Boolean
    ReDim blnLogUpToDate(LBound(FileName()) To UBound(FileName()))
    Dim logString As String
    Dim strTypeOfInstall As String

    Dim b As Long
    
    For b = LBound(FileName()) To UBound(FileName())
        
        ' Check if log dir/file exists, create if it doesn't, check last mod date if it does
        ' We don't need the true/false info for Installer, but we DO need to run these two
        ' functions to create directories if they don't exist yet
        
        ' If last mod date less than 1 day ago, CheckLog = True
        blnLogUpToDate(b) = CheckLog(strStyleDir(b), strLogDir(b), strFullLogPath(b))
        'Debug.Print FileName(b) & " log exists and was checked today: " & blnLogUpToDate(b)
        
        ' Check if template exists, if not create any missing directories
        blnTemplateExists(b) = IsTemplateThere(FinalDir(b), FileName(b), strFullLogPath(b))
        ' Debug.Print FileName(b) & " exists: " & blnTemplateExists(b)
        
        ' ===============================
        ' FOR DEBUGGING: SET TO TRUE,    |
        ' SO ALWAYS DOWNLOADS FILES      |
        ' Installer = True              '|
        ' ===============================
        
        If Installer = False Then 'Because if it's an installer, we just want to install the file

                
            ' ==========================================
            ' FOR DEBUGGING: SET TO FALSE AND THEN TRUE |
            ' TO TEST NEEDUPDATE FUNCTION               |
            ' blnLogUpToDate(b) = False                '|
            ' blnTemplateExists(b) = True              '|
            ' ==========================================
                
            If blnLogUpToDate(b) = True And blnTemplateExists(b) = True Then ' already checked today, already exists
                installCheck(b) = False
            ElseIf blnLogUpToDate(b) = False And blnTemplateExists(b) = True Then 'Log is new or not checked today, already exists
                'check version number
                installCheck(b) = NeedUpdate(Staging, FinalDir(b), FileName(b), strFullLogPath(b))
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
    If IsArrayEmpty(strInstallFile()) = True Then       ' No files need to be installed
        If Installer = True Then  ' Though this option (no files to install on installer) shouldn't actually occur
            #If Mac Then    ' because application.quit generates error on Mac
                ActiveDocument.Close (wdDoNotSaveChanges)
            #Else
                Application.Quit (wdDoNotSaveChanges)
            #End If
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
                ActiveDocument.Close (wdDoNotSaveChanges)
            End If
            
            Exit Sub
        End If
    End If
    
    ' ---------------- Close any open docs (with prompt) -----------------------------------
    Call CloseOpenDocs
        
    '----------------- download template files ------------------------------------------
    Dim d As Long
    
    For d = LBound(strInstallFile()) To UBound(strInstallFile())
        'If False, error in download; user was notified in DownloadFromConfluence function
        If DownloadFromConfluence(Staging, strInstallDir(d), strFullLogPath(d), strInstallFile(d)) = False Then
            If Installer = True Then
                #If Mac Then    ' because application.quit generates error on Mac
                    ActiveDocument.Close (wdDoNotSaveChanges)
                #Else
                    Application.Quit (wdDoNotSaveChanges)
                #End If
            Else
                Exit Sub
            End If
        End If
        
        ' If we just updated the main template, delete the old toolbar
        ' Will be added again by MacmillanGT AutoExec when it's launched, to capture updates
        #If Mac Then
            If strInstallFile(d) = "MacmillanGT.dotm" Then
                For Each Bar In CommandBars
                    If Bar.Name = "Macmillan Tools" Then
                        Bar.Delete
                        'Exit For  ' Actually don't exit, in case there are multiple toolbars
                    End If
                    Next
            End If
        #End If
    Next d
    
    '------Display installation complete message   ---------------------------
    Dim strComplete As String
    Dim strInstallType As String
    
    ' Quit if it's an installer, but not if it's an updater (updater was causing conflicts between GT and GtUpdater)
    If Installer = True Then
        strInstallType = "installed"
    Else
        strInstallType = "updated"
    End If
    
    strComplete = "The " & TemplateName & " has been " & strInstallType & " on your computer." & vbNewLine & vbNewLine & _
            "You must QUIT and RESTART Word for the changes to take effect."
    MsgBox strComplete, vbOKOnly, "Installation Successful"
    
    ' Mac 2011 Word can't do Application.Quit, so then just prompt user to restart and close Installer
    ' (but don't quit Word). Otherwise, quit for user on PC.
    ' Don't want to Close/Quit if it's an updater, because both MacmillanGT and GtUpdater need to run consecutively
    If Installer = True Then
        #If Mac Then
            ActiveDocument.Close (wdDoNotSaveChanges)
        #Else
            Application.Quit (wdDoNotSaveChanges)
        #End If
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

Private Function NeedUpdate(StagingURL As Boolean, Directory As String, FileName As String, Log As String) As Boolean
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
        
        #If Mac Then
            Call OpenDocMac(strFullTemplatePath)
        #Else
            Call OpenDocPC(strFullTemplatePath)
        #End If
        
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
    If DownloadFromConfluence(StagingURL, Directory, Log, strVersion) = False Then
        NeedUpdate = False
    End If
        
    '-------------------- Get version number of current template ---------------------
    If IsItThere(strFullVersionPath) = True Then
        NeedUpdate = True
        Dim strCurrentVersion As String
        strCurrentVersion = ImportVariable(strFullVersionPath)
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

Private Sub OpenDocMac(FilePath As String)
        Documents.Open FileName:=FilePath, ReadOnly:=True ', Visible:=False      'Mac Word 2011 doesn't allow Visible as an argument :(
End Sub

Private Sub OpenDocPC(FilePath As String)
        Documents.Open FileName:=FilePath, ReadOnly:=True, Visible:=False      'Win Word DOES allow Visible as an argument :)
End Sub



Private Function ImportVariable(strFile As String) As String
 
    Open strFile For Input As #1
    Line Input #1, ImportVariable
    Close #1
 
End Function



