Attribute VB_Name = "ViewStyles"
Option Explicit
Sub StylesViewLaunch()
' added by EW for v. 3.2
' runs different macros based on OS
' Set button and keyboard shortcut to run this macro

'Test the conditional compiler constant
    #If Mac Then
        'I am a Mac and will test if it is Word 2011 or higher
        If Val(Application.Version) > 14 Then
            Call Mac_BestStylesView
        End If
    #Else
        'I am Windows
        Call PC_BestStylesView
        #End If

End Sub
Private Sub PC_BestStylesView()
' added by EW for v. 3.2
' Setup for PC

Application.TaskPanes(wdTaskPaneFormatting).Visible = True          'Opens Styles Pane
Application.TaskPanes(wdTaskPaneStyleInspector).Visible = True     'Opens Style Inspector
ActiveDocument.FormattingShowFont = True                                     'Selects three center boxes in Styles Pane Options
ActiveDocument.FormattingShowParagraph = True
ActiveDocument.FormattingShowNumbering = True
ActiveDocument.FormattingShowFilter = wdShowFilterStylesAll         'Shows all styles
ActiveDocument.StyleSortMethod = wdStyleSortByName                     'Sorts styles alphabetically
ActiveDocument.ActiveWindow.View.ShowAll = True                          'Shows nonprinting characters and hidden text
ActiveDocument.ActiveWindow.View.Type = wdNormalView              'Switches to Normal/Draft view
ActiveDocument.ActiveWindow.StyleAreaWidth = InchesToPoints(1.5)                           'Sets Styles margin area in draft view to 1.5 in.

End Sub
Private Sub Mac_BestStylesView()
' added by EW for v. 3.2
' Setup for Mac

Application.Dialogs(1755).Display                                                       'opens the Styles Toolbox! Hurray!
ActiveDocument.ActiveWindow.View.ShowAll = True                         'Shows nonprinting characters and hidden text
ActiveDocument.ActiveWindow.View.Type = wdNormalView                'Switches to Normal/Draft view
ActiveDocument.ActiveWindow.StyleAreaWidth = InchesToPoints(1.5)                           'Sets Styles margin area in draft view to 1.5 in.
End Sub
