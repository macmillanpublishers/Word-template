Attribute VB_Name = "RibbonControl"
Option Explicit
Public myRibbon As IRibbonUI

Sub Onload(ribbon As IRibbonUI)
  'Creates a ribbon instance for use in this project
  Set myRibbon = ribbon
End Sub

Sub LaunchMacros(control As IRibbonControl)
    'Calls each macro when button is pressed on custom Ribbon
    Select Case control.ID
        Case Is = "BtnAttachTemplate"
            AttachTemplateMacro.zz_AttachStyleTemplate
        Case Is = "BtnCoverCopy"
            AttachTemplateMacro.zz_AttachCoverTemplate
        Case Is = "BtnRemoveColor"
            AttachTemplateMacro.zz_AttachBoundMSTemplate
        Case Is = "BtnViewStyles"
            ViewStyles.StylesViewLaunch
        Case Is = "BtnCastoff"
            CastoffMacro.UniversalCastoff
        Case Is = "BtnCleanup"
            CleanupMacro.MacmillanManuscriptCleanup
        Case Is = "BtnCharStyles"
            CharacterStyles.MacmillanCharStyles
        Case Is = "BtnStyleReport"
            Reports.MacmillanStyleReport
        Case Is = "BtnBkmkrCheck"
            Reports.BookmakerReqs
        Case Is = "BtnGtVersion"
            VersionCheck.CheckMacmillanGT
        Case Is = "BtnStyleVersion"
            VersionCheck.CheckMacmillan
        Case Is = "BtnLocTags"
            LOCtagsMacro.LibraryOfCongressTags
        Case Is = "BtnPrintStyles"
            PrintStyles.PrintStyles
        Case Is = "BtnTriceratops"
            EasterEggs.Triceratops
        'Case Is = "BtnEndnotes"
            'Endnotes.Unlink 'Add module and sub name here
        Case Else
            'Do nothing
    End Select
End Sub

Sub getLabel(control As IRibbonControl, ByRef returnedVal)
Dim TheOS As String
TheOS = System.OperatingSystem

    Select Case control.ID
        Case Is = "BtnViewStyles"
            If TheOS Like "*Mac*" Then
                If Application.Dialogs(1755).Show Then
                    returnedVal = "Hide Styles View"
                Else
                    returnedVal = "Show Styles View"
                End If
            Else
                If Application.TaskPanes(wdTaskPaneFormatting).Visible = True Then
                    returnedVal = "Hide Styles View"
                Else
                    returnedVal = "Show Styles View"
                End If
            End If
        End Select
End Sub

Sub buttonPressed(control As IRibbonControl, ByRef returnedVal)
Dim TheOS As String
TheOS = System.OperatingSystem

    Select Case control.ID
        Case Is = "TogViewStyles"
            If TheOS Like "*Mac*" Then
                If Application.Dialogs(1755).Show Then
                    returnedVal = True
                Else
                    returnedVal = False
                End If
            Else
                If Application.TaskPanes(wdTaskPaneFormatting).Visible = True Then
                    returnedVal = True
                Else
                    returnedVal = False
                End If
            End If
    End Select
End Sub
