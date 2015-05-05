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

'Open styles pane if closed, close if open
If Application.TaskPanes(wdTaskPaneFormatting).Visible = False Then
    Application.TaskPanes(wdTaskPaneFormatting).Visible = True
    
    'Selects three center boxes in Styles Pane Options, always
    ActiveDocument.FormattingShowFont = True
    ActiveDocument.FormattingShowParagraph = True
    ActiveDocument.FormattingShowNumbering = True

    'Shows all styles in document
    ActiveDocument.FormattingShowFilter = wdShowFilterStylesAvailable
    
    'Sorts styles alphabetically
    ActiveDocument.StyleSortMethod = wdStyleSortByName
Else
    Application.TaskPanes(wdTaskPaneFormatting).Visible = False
End If

'Open Style Inspector if closed, else close if open
If Application.TaskPanes(wdTaskPaneStyleInspector).Visible = False Then
    Application.TaskPanes(wdTaskPaneStyleInspector).Visible = True
Else
    Application.TaskPanes(wdTaskPaneStyleInspector).Visible = False
End If

'Shows nonprinting characters and hidden text if off, turns off if already on
If ActiveDocument.ActiveWindow.View.ShowAll = False Then
    ActiveDocument.ActiveWindow.View.ShowAll = True
Else
    ActiveDocument.ActiveWindow.View.ShowAll = False
End If

' Switches to Normal/Draft view if in Print, or vice versa
If ActiveDocument.ActiveWindow.View.Type = wdNormalView Then
    ActiveDocument.ActiveWindow.View.Type = wdPrintView
Else
    ActiveDocument.ActiveWindow.View.Type = wdNormalView
    ActiveDocument.ActiveWindow.StyleAreaWidth = InchesToPoints(1.5) 'Sets Styles margin area in draft view to 1.5 in
End If


End Sub
Private Sub Mac_BestStylesView()
' added by EW for v. 3.2
' Setup for Mac

'opens the Styles Toolbox! Hurray!
If Application.Dialogs(1755).Display = False Then
    Application.Dialogs(1755).Display
Else
    Application.Dialogs(1755).Hide
End If

'Shows nonprinting characters and hidden text
If ActiveDocument.ActiveWindow.View.ShowAll = False Then
    ActiveDocument.ActiveWindow.View.ShowAll = True
Else
    ActiveDocument.ActiveWindow.View.ShowAll = False
End If

'Switches to Normal/Draft view
If ActiveDocument.ActiveWindow.View.Type = wdNormalView Then
    ActiveDocument.ActiveWindow.View.Type = wdPrintView
Else
    ActiveDocument.ActiveWindow.View.Type = wdNormalView
    ActiveDocument.ActiveWindow.StyleAreaWidth = InchesToPoints(1.5)    'Sets Styles margin area in draft view to 1.5 in.
End If

                           
End Sub
