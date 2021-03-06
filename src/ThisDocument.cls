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

' ===== Document_Open ================================================================
' This sub will run automatically when a NEW document is created from this template,
' i.e. when a user double-clicks to run the installer.

Private Sub Document_New()

' Check if we're on the server; if so, don't run this check
  If WT_Settings.InstallType <> "user" Then
    DebugPrint "On server, quitting AutoExec"
  ' NOTE! Must use 'End' here because 'Exit' won't clear static variable (WT_Settings)
    End
  End If

' CHECK IF FLAG FILE IS PRESENT IN SAME DIR, IF YES THEN QUIT
' WILL BE ADDED WHEN WE'RE UNLOAD/LOADING ADDINS AS PART OF UPDATE/INSTALL STUFF
' AND DON'T WANT THIS TO LAUNCH EVERY TIME
  Dim blnDisableAutoMacroFileExists As Boolean
  blnDisableAutoMacroFileExists = IsItThere(ThisDocument.Path & _
    Application.PathSeparator & "DISABLE_AUTO_EXEC.txt")
  If blnDisableAutoMacroFileExists = True Then
    Exit Sub
  End If

'===== PART 2 ============================
  Dim blnInstaller As Boolean
  Dim strTemplateName As String
  Dim colTemplateNames As Collection

' True = standalone installer file
' False = daily update checker already installed
  blnInstaller = True

  'Plain English name of what you're installing/updating, should be singular
  strTemplateName = "Macmillan Tools and Style Template"

  'Get array of paths to STYLE TEMPLATE files to install
  Set colTemplateNames = SharedFileInstaller.GetTemplatesList(TemplatesYouWant:=installTemplates)
  
  ' Unload all add-ins first
  Dim strFileNames As Variant
  
  For Each strFileNames In colTemplateNames
    On Error Resume Next        'Error = add-in not available, don't need to uninstall
      AddIns(strFileNames).Installed = False
    On Error GoTo ErrHandler
  Next

' Install ALL template files (even if already exist)
  Call SharedFileInstaller.Installer(Installer:=blnInstaller, _
      TemplateName:=strTemplateName, TemplatesToInstall:=colTemplateNames)

' Load GtUpdater add-in (which will then load MacGT add-in)
  On Error Resume Next        'Error = add-in not available
    AddIns(colTemplateNames.Item(1)).Installed = True
  On Error GoTo 0

  ThisDocument.Close
  Exit Sub

ErrHandler:
  MsgBox Err.Number & ": " & Err.Description
End Sub


Sub AutoExec()
  DebugPrint "Word-template AutoExec is running"

' Check if we're on the server; if so, don't run this check
  If WT_Settings.InstallType <> "user" Then
    DebugPrint "On server, quitting AutoExec"
  ' NOTE! Must use 'End' here because 'Exit' won't clear static variable (WT_Settings)
    End
  Else
'    DebugPrint "Install is: " & WT_Settings.InstallType
  End If

' CHECK IF FLAG FILE IS PRESENT IN SAME DIR, IF YES THEN QUIT
' WILL BE ADDED WHEN WE'RE UNLOAD/LOADING ADDINS AS PART OF UPDATE/INSTALL STUFF
' AND DON'T WANT THIS TO LAUNCH EVERY TIME
  Dim blnDisableAutoMacroFileExists As Boolean
  blnDisableAutoMacroFileExists = IsItThere(ThisDocument.Path & _
    Application.PathSeparator & "DISABLE_AUTO_EXEC.txt")
  If blnDisableAutoMacroFileExists = True Then
    Exit Sub
  End If

    ' Check if toolbar is present if Mac, if not add it.
    ' If PC, remove "Add-Ins" tab that gets added for some reason.
    Dim Bar As CommandBar
    #If Mac Then
        Dim blnToolbar As Boolean
        For Each Bar In CommandBars
            If Bar.Name = "Macmillan Tools" Then
                blnToolbar = True
                Exit For
            Else
                blnToolbar = False
            End If
        Next

        If blnToolbar = False Then
            Application.StatusBar = "Updating Macmillan Tools toolbar"
            Call CreateMacToolbar(PathToTemplate:=ThisDocument.Path)
        End If
    #Else
        ' If opened read-only, can't save later
        If IsReadOnly(Path:=ThisDocument.FullName) = False Then
            For Each Bar In CommandBars
                If Bar.Name = "Macmillan Tools" Then
                    Bar.Delete
                    Exit For
                End If
            Next
            ' Need to save template or user is prompted to save when closing Word
            ThisDocument.Save
        End If
    #End If


'===== PART 2 ============================
  Dim blnInstaller As Boolean
  Dim strTemplateName As String
  Dim colTemplateNames As Collection

' True = standalone installer file
' False = daily update checker already installed
  blnInstaller = False

  'Plain English name of what you're installing/updating, should be singular
  strTemplateName = "Macmillan Styles Template"

  'Get array of paths to STYLE TEMPLATE files to install
  Set colTemplateNames = SharedFileInstaller.GetTemplatesList(TemplatesYouWant:=stylesTemplates)

' CHECK/UPDATE STYLES TEMPLATES ONLY.
  Call SharedFileInstaller.Installer(Installer:=blnInstaller, _
      TemplateName:=strTemplateName, TemplatesToInstall:=colTemplateNames)

  ' BLAH BLAH ONTIME NOT WORKING JUST DON'T CHECK ANY MORE WE'LL GET RID OF WITH CONFIG FILE.
  ' Now check GtUpdater.dotm, which can't have code running when we check it, so have to wait 1 second for GtUpdater to
  ' finish running (because this AutoExec is called from GtUpdater in Startup
'    Application.OnTime When:=Now + TimeValue("00:00:01"), _
      Name:=ThisDocument.CheckGtUpdater(Branch:=downloadBranch, InstallerOrNot:=blnInstaller)

'    Call ThisDocument.CheckGtUpdater(InstallerOrNot:=blnInstaller)

  On Error GoTo 0
  Exit Sub

ErrHandler:
    MsgBox Err.Number & ": " & Err.Description
End Sub


Function CheckGtUpdater(InstallerOrNot As Boolean)

' Checks version number of GtUpdater.dotm once a day and updates to new version if necessary
' GtUpdater.dotm can't be running code when this launches.

'===== DEPENDENCIES ========================
' This is part 1 of 2. It calls the Installer sub, which downloads and installs the Macmillan style templates and macros.
' If this is meant to be an installer, it should be called "Document_Open()" and reside in the ThisDocument module in a .docm file.
' If this is meant to be an updater, it should be called "AutoExec()" and reside in the ThisDocument module in a .dotm file,
' and be saved in the STARTUP directory.

  DebugPrint "CheckGtUpdater is running."

  Dim strTemplateName As String
  strTemplateName = "Macmillan Updater"

  ' get array of files to install
  Dim colInstallMe As Collection
  Set colInstallMe = GetTemplatesList(TemplatesYouWant:=updaterTemplates)

  ' Installer = True if this is for a standalone installation file;
  ' False is for a daily update checker
  Call SharedFileInstaller.Installer(Installer:=InstallerOrNot, _
      TemplateName:=strTemplateName, TemplatesToInstall:=colInstallMe)

  ' ====== PART 2: LOAD TEMPLATE AGAIN ==============
  If IsItThere(colInstallMe.Item(1)) = True Then
  ' Have to add Template before loading if not in Startup
  ' Could do Install:=True with Add method, but it somehow broke everything
    AddIns.Add (colInstallMe.Item(1))
    AddIns(colInstallMe.Item(1)).Installed = True
  Else
    Dim strMessage As String
      strMessage = "Wow, this is really embarrassing but something went wrong with your Macmillan Updater template. Please send an email to " & _
      "workflows@macmillan.com and let them know that something is broken!"
    MsgBox strMessage, vbOKOnly, "We're So Sorry"
  End If

  On Error GoTo 0
  Exit Function

ErrHandler:
  MsgBox Err.Number & ": " & Err.Description
End Function


Private Sub CreateMacToolbar(PathToTemplate As String)
' ====== USE ======
' Creates custom toolbar on a Mac. Don't want to do it manually because saving on Mac
' removes the custom PC Ribbon
' See http://www.fontstuff.com/ebooks/free/fsexceladdins.pdf
' And http://word.mvps.org/faqs/macrosvba/SetCustomButtonImage.htm
'
' ====== DEPENDENCIES =======
' Obviously, custom ribbon XML code must have been added to the template on PC
' See XML code in this repo for format. Macro name MUST be stored as button ID attribute.
' CANNOT OPEN AND SAVE TEMPLATE ON MAC 2011! This removes the customUI directory.
' When you create the custom ribbon on PC, must include Mac button images as well,
' which should have the same file name but ending in "_mac" and be PNG files
' that are 16 x 16 pixels, background RGB 191, 191, 191.

    Dim Bar As CommandBar
    Application.ScreenUpdating = False

    #If Mac Then
        ' Code below is all Mac specific (paths and shell code) but could probably expand
        ' to PC if we ever want to support Word versions earlier than 2007 (no ribbon).
        ' Would have to figure out how to send shell commands on PC tho

'        Dim StartTime As Double
'        Dim SecondsElapsed As Double
'
'        'Remember time when macro starts
'        StartTime = Timer
        '---------------------------------------------

        Dim strPath As String
        Dim strFile As String
        Dim strMacZip As String
        Dim strMacUnzip As String
        Dim strMacTmp As String
        Dim strBashUnzip As String
        Dim strBashZip As String
        Dim strMacXML As String
        Dim strXmlPath As String
        Dim strImagePath As String
        Dim strMacImagePath As String

        MsgBox "Please click OK and then WAIT a few seconds while we create " & _
            "the Macmillan Tools toolbar for you." & vbNewLine & vbNewLine & _
            "A notice will pop up when it's complete!"

        ' Set up all our paths to files/directories
        strPath = PathToTemplate
        
        ' Get just the file name w/o path or extension
        strFile = Mid(strPath, InStrRev(strPath, ":") + 1, InStrRev(strPath, ".") - InStrRev(strPath, ":") - 1)
        
        'Stop
        

        ' location we're going to copy the template to as a zip file
        strMacTmp = MacScript("path to temporary items as string")
        'DebugPrint "strMacTmp: " & strMacTmp
        strMacUnzip = strMacTmp & strFile
        'DebugPrint "strMacUnzip: " & strMacUnzip
        strMacZip = strMacUnzip & ".zip"
        'DebugPrint "strMacZip: " & strMacZip
        strMacXML = strMacUnzip & ":customUI:"
        'DebugPrint "strMacXML: " & strMacXML
        strMacImagePath = strMacXML & "images:"
        'DebugPrint "strMacImagePath: " & strMacImagePath
        ' convert formats of above for command line

        strBashUnzip = Replace(Right(strMacUnzip, Len(strMacUnzip) - (InStr(strMacUnzip, ":") - 1)), ":", "/")
        'DebugPrint "strBashUnzip: " & strBashUnzip
        strBashZip = Replace(Right(strMacZip, Len(strMacZip) - (InStr(strMacZip, ":") - 1)), ":", "/")
        'DebugPrint "strBashZip: " & strBashZip
        strXmlPath = Replace(Right(strMacXML, Len(strMacXML) - (InStr(strMacXML, ":") - 1)), ":", "/")
        'DebugPrint "strXmlPath: " & strXmlPath
        ' path to button images
        strImagePath = Replace(Right(strMacImagePath, Len(strMacImagePath) - (InStr(strMacImagePath, ":") - 1)), ":", "/")
        'DebugPrint "strImagePath: " & strImagePath

        ' ===== First we copy the template to tmp as a .zip and then unzip it ====='
        ' Copy this template to tmp as a zip file
        FileCopy strPath, strMacZip

        ' Delete any previous unzipped folders in tmp
        ' Unzip the copy we just made, delete the zip file
        ' -u means update, i.e. overwrite any files that currently exist
        ShellAndWaitMac ("rm -r " & strBashUnzip & ";unzip -u " & strBashZip & " -d " & strBashUnzip & ";rm " & strBashZip)

        ' ===== Then we build a toolbar from the values in the customUI.xml file ===== '
        ' To edit the shell commands read up on xmllint and xpath:
        ' http://blog.powered-up-games.com/wordpress/archives/70'
        ' https://docs.oracle.com/javase/tutorial/jaxp/xslt/xpath.html
        ' http://www.thegeekstuff.com/2014/12/linux-parse-xml/


        Dim strCmdStart As String
        Dim strCmdEnd As String
        Dim strXpath As String
        Dim lngTabCount As Long
        Dim strTabName As String
        Dim lngGroupCount As Long
        Dim lngButtonCount As Long
        Dim A As Long
        Dim B As Long
        Dim C As Long
        Dim h As Long
        Dim NewToolbar As CommandBar

        Dim strImageName As String
        Dim blnBeginGroup As Boolean
        Dim strCaption As String
        Dim strAction As String
        Dim shpButtonPic As Shape
        Dim NewButton As CommandBarButton


        If IsItThere(strMacXML) = True Then
            ' first part "sed -e "s/xmlns/ignore/" customUI.xml" removes namespace, which messes up xmllint
            ' that is then piped to xmllint which gets the attribute value based on the Xpath
            strCmdStart = "sed -e \" & Chr(34) & "s/xmlns/ignore/\" & Chr(34) & " " & strXmlPath & "customUI.xml | xmllint --xpath \" & Chr(34)
            strCmdEnd = "\" & Chr(34) & " -"

            ' Get count of custom tabs in XML
            strXpath = "count(//tab)"
            lngTabCount = ShellAndWaitMac(strCmdStart & strXpath & strCmdEnd)

            ' Get name of toolbar from XML
            If lngTabCount > 0 Then
                For A = 1 To lngTabCount
                    strXpath = "string(//tab[" & A & "]/@label)"
                    strTabName = ShellAndWaitMac(strCmdStart & strXpath & strCmdEnd)

                    ' Test if toolbar already exists, and delete it - in case there are updates
                    For Each Bar In CommandBars
                        If Bar.Name = strTabName Then
                            Bar.Delete
                            'Exit For
                        End If
                    Next

                    'Stop        ' For testing delete loop above

                    ' Create custom toolbar
                    Set NewToolbar = Application.CommandBars.Add(Name:=strTabName, Temporary:=False)

                    ' Count number of groups in tab
                    strXpath = "count(//tab[" & A & "]/group)"
                    lngGroupCount = ShellAndWaitMac(strCmdStart & strXpath & strCmdEnd)

                    If lngGroupCount > 0 Then
                        For B = 1 To lngGroupCount
                            ' Count number of tabs in group
                            strXpath = "count(//tab[" & A & "]/group[" & B & "]/button)"
                            lngButtonCount = ShellAndWaitMac(strCmdStart & strXpath & strCmdEnd)

                            If lngButtonCount > 0 Then

                                ' Add new buttons
                                For C = 1 To lngButtonCount

                                    ' Start of Xpath for all button attributes - need to add attribute name AND close parens to use
                                    strXpath = "string(//tab[" & A & "]/group[" & B & "]/button[" & C & "]/@"

                                    ' If first button in group, set BeginGroup to True
                                    If C = 1 Then blnBeginGroup = True Else blnBeginGroup = False

                                    ' Get image name
                                    strImageName = ShellAndWaitMac(strCmdStart & strXpath & "image)" & strCmdEnd)
                                    strImageName = strMacImagePath & strImageName & "_mac.png"

                                    ' Get button caption and macro name
                                    strCaption = ShellAndWaitMac(strCmdStart & strXpath & "label)" & strCmdEnd)
                                    strAction = ShellAndWaitMac(strCmdStart & strXpath & "id)" & strCmdEnd)

                                    ' Copy button image to clipboard w/ AppleScript
                                    Dim strCommand As String
                                    ' 199 is << and 200 is >> (but as a single character)
                                    strCommand = "set buttonPic to read file " & Chr(34) & strImageName & Chr(34) & " as " & Chr(199) & _
                                        "class PNGf" & Chr(200) & Chr(13) & "set the clipboard to buttonPic"
                                    'DebugPrint strCommand
                                    MacScript (strCommand)

                                    ' Create button
                                    ' This includes an icon and caption for each button
                                    Set NewButton = NewToolbar.Controls.Add(Type:=msoControlButton)
                                    With NewButton
                                        .PasteFace
                                        .Caption = strCaption
                                        .Style = msoButtonIconAndCaption
                                        .OnAction = strAction
                                        .BeginGroup = blnBeginGroup
                                    End With
                                Next C
                            End If
                        Next B
                    End If
                Next A
            End If

            ' Set Toolbar dimension AFTER creation
            With NewToolbar
                .Height = 252 'pixels, I assume
                .Width = 171
                .Visible = True
            End With

            ' Delete the unzipped directory
            ShellAndWaitMac ("rm -rf " & strBashUnzip)
        Else
            'There is no customUI dir in template
            Dim strMessage As String
            strMessage = "The Macmillan Tools template cannot create the macro toolbar. Please contact " & _
                "workflows@macmillan.com for assistance."

            MsgBox strMessage, vbCritical, "Toolbar Error"

        End If

'        '----------------------Timer End-------------------------------------------
'        'Determine how many seconds code took to run
'        SecondsElapsed = Round(Timer - StartTime, 2)
'
'        'Notify user in seconds
'        MsgBox "This code ran successfully in " & SecondsElapsed & " seconds"
'        '---------------------------------------------------------------------------
        
        MsgBox "The Macmillan Tools toolbar has been updated!"
    
    #End If

End Sub


