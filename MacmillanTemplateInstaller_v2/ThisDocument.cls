VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'=======================================
' By Erica Warren - erica.warren@macmillan.com
'
'===== USE ==============================
' Sets variable names (file names and directories) for Installer sub in FileInstaller module
'
'===== DEPENDENCIES =====================
' This is part 1 of 2. It calls the Installer sub, which downloads and installs the Macmillan style templates and macros.
' If this is meant to be an installer, it should reside in the ThisDocument module in a .docm file.
' If this is meant to be an updater, it should reside in whatever module calls the checking procedure.

' Requires FileInstaller and SharedMacros modules be installed in the same template

Option Explicit
Option Base 1

Private Sub Document_Open()
    ' ==================================================
    ' For debugging; comment out of production code
    MsgBox "I am Doc_Open for Installer and I am running"
    ' ==================================================

    Dim blnStaging As Boolean
    Dim blnInstaller As Boolean
    Dim strTemplateName As String
    Dim strFileName() As String
    Dim strFileDir() As String
    
    Dim strPcStyleDir As String
    Dim strPcStartupDir As String
    Dim strMacStyleDir As String
    Dim strMacStartupDir As String
    Dim strMacDocs As String
    
'==========================================================================================
    'DEFINE ALL YOUR VARIABLES HERE
    
    'Make sure to switch to FALSE before going live!
    blnStaging = True
    
    'True = standalone installer file, False = daily update checker already installed
    blnInstaller = True
    
    'Plain English name of what you're installing/updating, should be singular
    strTemplateName = "Macmillan Tools and Style Template"
    
    'Update array lengths and template file names here
    ReDim strFileName(1 To 5)               'Declare number of files here
    ReDim strFileDir(1 To 5)                'And here
    
    'File names you're installing/updating
    strFileName(1) = "GtUpdater.dotm"
    strFileName(2) = "macmillan.dotm"
    strFileName(3) = "macmillan_NoColor.dotm"
    strFileName(4) = "MacmillanCoverCopy.dotm"
    strFileName(5) = "MacmillanGT.dotm"
    
    ' Enter the home directories for your templates here, for Mac *AND* PC
    #If Mac Then
        strMacDocs = MacScript("return (path to documents folder) as string")
        strMacStyleDir = strMacDocs & "MacmillanStyleTemplate"
        strMacStartupDir = Application.StartupPath
    #Else
        strPcStyleDir = Environ("PROGRAMDATA") & "\MacmillanStyleTemplate"
        strPcStartupDir = Application.StartupPath
    #End If
'=============================================================================================

    
    #If Mac Then
        strFileDir(1) = strMacStartupDir      'Final directory you're installing file in on a Mac
        strFileDir(2) = strMacStyleDir
        strFileDir(3) = strMacStyleDir
        strFileDir(4) = strMacStyleDir
        strFileDir(5) = strMacStyleDir
    #Else
        strFileDir(1) = strPcStartupDir      'Final directory you're installing file in on a PC
        strFileDir(2) = strPcStyleDir
        strFileDir(3) = strPcStyleDir
        strFileDir(4) = strPcStyleDir
        strFileDir(5) = strPcStyleDir
    #End If
        
    'Check that same number of file names and directories
    If UBound(strFileName()) <> UBound(strFileDir()) Then
        MsgBox "Number of files for download must match number of directories.", vbCritical, "OOPS!"
        Exit Sub
    End If

    ' ====== PART 1: REMOVE LEGACY TOOLS TEMPLATE =======
    ' The main template, MacmillanGT.dotm, used to be saved in Startup but since non-admin users
    ' don't have write access to that folder on Mac, we've moved it and will load from this sub. But need to
    ' delete the old one first
    Dim strOldGT As String

    #If Mac Then
        strOldGT = strMacStartupDir & Application.PathSeparator & strFileName(5)
    #Else
        strOldGT = strPcStartupDir & Application.PathSeparator & strFileName(5)
    #End If

    On Error GoTo ErrHandler
    If IsItThere(strOldGT) = True Then
        AddIns(strOldGT).Installed = False
        Kill strOldGT
    End If

    ' ====== PART 2: CHECK FOR UPDATES ===============
    ' Unload all add-ins first
    Dim strTemplate As String
    Dim a As Long
    
    For a = LBound(strFileName()) To UBound(strFileName())
        strTemplate = strFileDir(a) & Application.PathSeparator & strFileName(a)
        On Error Resume Next        'Error = add-in not available, don't need to uninstall
            AddIns(strTemplate).Installed = False
        On Error GoTo 0
    Next a
    
    ' Run Installer macro
    'Installer = True if this is for a standalone installation file; False is for a daily update checker
    Call SharedFileInstaller.Installer(Staging:=blnStaging, Installer:=blnInstaller, TemplateName:=strTemplateName, FileName:=strFileName(), FinalDir:=strFileDir())
    
    ' Mac 2011 Word can't do Application.Quit, so then just prompt user to restart and close Installer
    ' (but don't quit Word). Otherwise, quit for user on PC.
    '#If Mac Then
        On Error GoTo 0
        Application.Quit (wdDoNotSaveChanges)
    '#Else
        'Application.Quit (wdDoNotSaveChanges)
    '#End If
    
    'On Error GoTo 0
    'Exit Sub
    
ErrHandler:
    If Err.Number = 75 Then     ' Change to whatever the actual error is when try to kill file
        Dim strMessage As String
        ' Can we tell if this is a Macmillan install, and prompt them to use Self Service?"
        ' Could we launch Self Service from AppleScript or something?
        strMessage = "Uh oh, looks like you need to re-install the Macmillan template from Self Service or use an account with administrator priviledges."
        MsgBox strMessage, vbOKOnly, "New template needed"
    Else
        Debug.Print Err.Number & ": " & Err.Description
    End If
    
End Sub


