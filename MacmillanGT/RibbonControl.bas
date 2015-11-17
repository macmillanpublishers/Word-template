Attribute VB_Name = "RibbonControl"
Option Explicit
Public myRibbon As IRibbonUI

Sub Onload(ribbon As IRibbonUI)
  'Creates a ribbon instance for use in this project
  Set myRibbon = ribbon
End Sub

Sub LaunchMacros(Optional control As IRibbonControl, Optional buttonID As String)
    ' Mac launches from custom toolbar, not ribbonControl
    #If Mac Then
        Select Case buttonID
    #Else
    'Calls each macro when button is pressed on custom Ribbon
        Select Case control.ID
    #End If
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
            Case Is = "BtnEndnotes"
                Endnotes.EndnoteDeEmbed
            Case Else
                'Do nothing
        End Select
End Sub
