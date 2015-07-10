VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Sub AutoExec()
'=======================================
' By Erica Warren - erica.warren@macmillan.com
'
'===== USE ===============================
' Sets variable names (file names and directories) for Installer sub in FileInstaller module.
'
'===== DEPENDENCIES ========================
' This is part 1 of 2. It calls the Installer sub, which downloads and installs the Macmillan style templates and macros.
' If this is meant to be an installer, it should be called "Document_Open()" and reside in the ThisDocument module in a .docm file.
' If this is meant to be an updater, it should be called "AutoExec()" and reside in the ThisDocument module in a .dotm file,
' and be saved in the STARTUP directory.
        
    ' For testing / debugging
    'MsgBox "I am GtUpdater and I am running!"
            
    Dim strTemplateName As String
    Dim strFileName() As String
    Dim strFileDir() As String
    
    Dim strPcStyleDir As String
    Dim strPcStartupDir As String
    Dim strMacStyleDir As String
    Dim strMacStartupDir As String
    Dim strMacUser As String
    
    ' ==========================================================================================
    'DEFINE ALL YOUR VARIABLES HERE
    
    ' Enter the home directories for your templates here, for Mac *AND* PC
    
    #If Mac Then
        strMacUser = MacScript("tell application " & Chr(34) & "System Events" & Chr(34) & Chr(13) & _
                "return (name of current user)" & Chr(13) & "end tell")
        strMacStyleDir = "Macintosh HD:Users:" & strMacUser & ":Documents:MacmillanStyleTemplate"
        strMacStartupDir = "Macintosh HD:Applications:Microsoft Office 2011:Office:Startup:Word"
    #Else
        strPcStyleDir = Environ("PROGRAMDATA") & "\MacmillanStyleTemplate"
        strPcStartupDir = Environ("APPDATA") & "\Microsoft\Word\STARTUP"
    #End If
    
    'Update array lengths and template file names here
    
    strTemplateName = "Macmillan Tools and Style Template"           ''Plain English name of what you're installing/updating, should be singular
    
    ReDim strFileName(1 To 1)               'Declare number of files here
    ReDim strFileDir(1 To 1)                'And here
    
    'File 1
    strFileName(1) = "MacmillanGT2.dotm"                                 'File name you're installing/updating
    'strFileName(2) = "macmillan.dotm"
    'strFileName(3) = "macmillan_NoColor.dotm"
    'strFileName(4) = "MacmillanCoverCopy.dotm"
    'strFileName(5) = "GtUpdater.dotm"
    
    '=============================================================================================
    
    #If Mac Then
        strFileDir(1) = strMacStartupDir      'Final directory you're installing file in on a Mac
        'strFileDir(2) = strMacStyleDir
        'strFileDir(3) = strMacStyleDir
        'strFileDir(4) = strMacStyleDir
        'strFileDir(5) = strMacStartupDir
    #Else
        strFileDir(1) = strPcStartupDir      'Final directory you're installing file in on a PC
        'strFileDir(2) = strPcStyleDir
        'strFileDir(3) = strPcStyleDir
        'strFileDir(4) = strPcStyleDir
        'strFileDir(5) = strPcStartupDir
    #End If
        
    'Check that same number of file names and directories
    If UBound(strFileName()) <> UBound(strFileDir()) Then
        MsgBox "Number of files for download must match number of directories.", vbCritical, "OOPS!"
        Exit Sub
    End If
    
    'Installer = True if this is for a standalone installation file; False is for a daily update checker
    Call FileInstaller.Installer(Installer:=False, TemplateName:=strTemplateName, FileName:=strFileName(), FinalDir:=strFileDir())
    
End Sub


