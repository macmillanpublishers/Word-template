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
'
' ********** PART ONE ************
'
' ====== USE ======
' Creates custom toolbar on a Mac. Don't want to do it manually because saving on Mac
' removes the custom PC Ribbon
' See http://www.fontstuff.com/ebooks/free/fsexceladdins.pdf
' And http://word.mvps.org/faqs/macrosvba/SetCustomButtonImage.htm
'
' ====== DEPENDENCIES =======
' Obviously, custom ribbon XML code must have been added to the template on PC
' CANNOT OPEN AND SAVE TEMPLATE ON MAC 2011! This removes the customUI directory
' When you create the custom ribbon on PC, must include Mac button images as well,
' which should have the same file name but ending in "_mac" and be PNG files

    #If Mac Then
        Dim strPath As String
        Dim strFileName As String
        Dim strZipPath As String
        Dim strUnzipPath As String

        ' ===== First we copy the template to tmp as a .zip and then unzip it ====='
        strPath = ThisDocument.Path
        'strPath = "Macintosh HD:Applications:Microsoft Office 2011:Office:Startup:Word:MacmillanGT.dotm"
        ''Debug.Print strPath

        strFileName = Mid(strPath, InStrRev(strPath, ":") + 1, InStrRev(strPath, ".") - InStrRev(strPath, ":") - 1)
        ''Debug.Print strFileName

        strZipPath = "Macintosh HD:private:tmp:" & strFileName & ".zip"
        ''Debug.Print strZipPath

        FileCopy strPath, strZipPath

        strUnzipPath = "/tmp/" & strFileName
        ''Debug.Print strUnzipPath

        strZipPath = strUnzipPath & ".zip"
        ''Debug.Print strZipPath

        ShellAndWaitMac ("unzip -u " & strZipPath & " -d " & strUnzipPath)

        ' ===== Then we build a toolbar from the values in the customUI.xml file ===== '
        ' To edit the shell commands read up on xmllint and xpath:
        ' http://blog.powered-up-games.com/wordpress/archives/70'
        ' https://docs.oracle.com/javase/tutorial/jaxp/xslt/xpath.html
        ' http://www.thegeekstuff.com/2014/12/linux-parse-xml/

        Dim strXmlPath As String
        Dim strCmdStart As String
        Dim strCmdEnd As String
        Dim strXpath As String
        Dim lngTabCount As Long
        Dim strTabName As String
        Dim lngGroupCount As Long
        Dim lngButtonCount As Long
        Dim a As Long
        Dim b As Long
        Dim c As Long
        Dim Bar As CommandBar
        Dim h As Long
        Dim NewToolbar As CommandBar
        Dim strImagePath As String
        Dim strImageName As String
        Dim blnBeginGroup As Boolean
        Dim strCaption As String
        Dim strAction As String

        strXmlPath = strUnzipPath & "/customUI/"
        ' path to button images
        strImagePath = strXmlPath & "images/"

        If IsItThere("Macintosh HD:private" & Replace(strXmlPath, "/", ":")) = True Then
            strCmdStart = "sed -e \" & Chr(34) & "s/xmlns/ignore/\" & Chr(34) & " " & strXmlPath & "customUI.xml | xmllint --xpath \" & Chr(34)
            strCmdEnd = "\" & Chr(34) & " -"

            ' Get count of custom tabs in XML
            strXpath = "count(//tab)"
            lngTabCount = ShellAndWaitMac(strCmdStart & strXpath & strCmdEnd)
            'Debug.Print "There are " & lngTabCount & " tab(s)"

            ' Get name of toolbar from XML
            If lngTabCount > 0 Then
                For a = 1 To lngTabCount
                    strXpath = "string(//tab[" & a & "]/@label)"
                    strTabName = ShellAndWaitMac(strCmdStart & strXpath & strCmdEnd)
                    'Debug.Print "Tab " & a & " is named " & strTabName

                    ' Test if toolbar already exists, and delete it
                    ' We'll re-create it each session, to be sure to capture updates
                    h = 0
                    For Each Bar In CommandBars
                        h = h + 1
                        'Debug.Print h & ": " & bar.Name
                        If Bar.Name = strTabName Then
                            Bar.Delete
                            Exit For
                        End If
                    Next

                    'Stop        ' For testing delete loop above

                    ' Create custom toolbar
                    Set NewToolbar = Application.CommandBars.Add(Name:=strTabName, Temporary:=True)

                    ' Count number of groups in tab
                    strXpath = "count(//tab[" & a & "]/group)"
                    lngGroupCount = ShellAndWaitMac(strCmdStart & strXpath & strCmdEnd)
                    'Debug.Print "There are " & lngGroupCount & " group(s) in tab " & a

                    If lngGroupCount > 0 Then
                        For b = 1 To lngGroupCount
                            ' Count number of tabs in group
                            strXpath = "count(//tab[" & a & "]/group[" & b & "]/button)"
                            lngButtonCount = ShellAndWaitMac(strCmdStart & strXpath & strCmdEnd)
                            'Debug.Print "There are " & lngButtonCount & " buttons in group " & b & " on tab " & a

                            If lngButtonCount > 0 Then

                                ' Add new buttons
                                For c = 1 To lngButtonCount

                                    ' Start of Xpath for all button attributes - need to add attribute name and close parens to use
                                    strXpath = "string(//tab[" & a & "]/group[" & b & "]/button[" & c & "]/@"

                                    ' If first button in group, set BeginGroup to True
                                    If c = 1 Then blnBeginGroup = True Else blnBeginGroup = False

                                    ' Get image name
                                    strImageName = ShellAndWaitMac(strCmdStart & strXpath & "image)" & strCmdEnd)
                                    'Debug.Print "This button's image file is called " & strImageName
                                    strImageName = "Macintosh HD:private" & Replace(strImagePath & strImageName & "_mac.png", "/", ":")
                                    'Debug.Print "This button's image is " & strImageName

                                    ' Add button image, copy to clipboard, delete
                                    Set shpButtonPic = ThisDocument.Shapes.AddPicture(FileName:=strImageName)
                                    shpButtonPic.Select
                                    Selection.CopyAsPicture
                                    Selection.Delete

                                    ' Get button caption and macro name
                                    strCaption = ShellAndWaitMac(strCmdStart & strXpath & "label)" & strCmdEnd)
                                    'Debug.Print "This button's caption is " & strCaption
                                    strAction = ShellAndWaitMac(strCmdStart & strXpath & "onAction)" & strCmdEnd)
                                    'Debug.Print "This button's macro is " & strAction

                                    ' Create button
                                    Set NewButton = NewToolbar.Controls.Add(Type:=msoControlButton)
                                    With NewButton
                                        .PasteFace
                                        .Caption = strCaption
                                        .Style = msoButtonIconAndCaption
                                        .OnAction = strAction
                                        .BeginGroup = blnBeginGroup
                                    End With
                                Next c
                            End If
                        Next b
                    End If
                Next a
            End If
            ' Don't need to delete the unzipped folder because the -u option in unzip overwrites any files

            ' Set Toolbar dimension
            With NewToolbar
                .Height = 252 'pixels, I assume
                .Width = 171
                .Visible = True
            End With

        Else
            'There is no customUI dir in template
            'Do some error handling/message to user
        End If

    #End If
    
    
' For testing part 1
' Stop


' ********* PART TWO ***************************


'===== USE ===============================
' Sets variable names (file names and directories) for Installer sub in FileInstaller module.
'
'===== DEPENDENCIES ========================
' This is part 1 of 2. It calls the Installer sub, which downloads and installs the Macmillan style templates and macros.
' If this is meant to be an installer, it should be called "Document_Open()" and reside in the ThisDocument module in a .docm file.
' If this is meant to be an updater, it should be called "AutoExec()" and reside in the ThisDocument module in a .dotm file,
' and be saved in the STARTUP directory.
'
' Requires FileInstaller and SharedMacros modules be installed in the same template.
        
    ' ======= For testing / debugging ==============
    ' MsgBox "I am MacmillanGT and I am running!"
    ' ==============================================
    
    Dim blnStaging As Boolean
    Dim blnInstaller As Boolean
    Dim strTemplateName As String
    Dim strFileName() As String
    Dim strFileDir() As String
    
    Dim strPcStyleDir As String
    Dim strPcStartupDir As String
    Dim strMacStyleDir As String
    Dim strMacStartupDir As String
    Dim strMacUser As String
    
'==========================================================================================
    'DEFINE ALL YOUR VARIABLES HERE
    
    'Make sure to switch to FALSE before going live!
    blnStaging = False
    
    'True = standalone installer file, False = daily update checker already installed
    blnInstaller = False
    
    'Plain English name of what you're installing/updating, should be singular
    strTemplateName = "Macmillan Styles Template"
    
    'Update array lengths and template file names here
    ReDim strFileName(1 To 4)               'Declare number of files here
    ReDim strFileDir(1 To 4)                'And here
    
    'File names you're installing/updating
    'strFileName(1) = "MacmillanGT.dotm"
    strFileName(1) = "GtUpdater.dotm"
    strFileName(2) = "macmillan.dotm"
    strFileName(3) = "macmillan_NoColor.dotm"
    strFileName(4) = "MacmillanCoverCopy.dotm"
    
    ' Enter the home directories for your templates here, for Mac *AND* PC
    #If Mac Then
        strMacUser = MacScript("tell application " & Chr(34) & "System Events" & Chr(34) & Chr(13) & _
                "return (name of current user)" & Chr(13) & "end tell")
        strMacStyleDir = "Macintosh HD:Users:" & strMacUser & ":Documents:MacmillanStyleTemplate"
        strMacStartupDir = Application.StartupPath
    #Else
        strPcStyleDir = Environ("PROGRAMDATA") & "\MacmillanStyleTemplate"
        strPcStartupDir = Application.StartupPath
    #End If
'=============================================================================================
    
    #If Mac Then
        'strFileDir(1) = strMacStartupDir      'Final directory you're installing file in on a Mac
        strFileDir(1) = strMacStartupDir
        strFileDir(2) = strMacStyleDir
        strFileDir(3) = strMacStyleDir
        strFileDir(4) = strMacStyleDir
    #Else
        'strFileDir(1) = strPcStartupDir      'Final directory you're installing file in on a PC
        strFileDir(1) = strPcStartupDir
        strFileDir(2) = strPcStyleDir
        strFileDir(3) = strPcStyleDir
        strFileDir(4) = strPcStyleDir
    #End If
        
    'Check that same number of file names and directories
    If UBound(strFileName()) <> UBound(strFileDir()) Then
        MsgBox "Number of files for download must match number of directories.", vbCritical, "OOPS!"
        Exit Sub
    End If
    
    'Installer = True if this is for a standalone installation file; False is for a daily update checker
    Call SharedFileInstaller.Installer(Staging:=blnStaging, Installer:=blnInstaller, TemplateName:=strTemplateName, FileName:=strFileName(), FinalDir:=strFileDir())
    
End Sub



