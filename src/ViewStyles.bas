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
    
    'in print view, toggle to draft view and open styles stuff
    If ActiveDocument.ActiveWindow.View.Type = wdPrintView Then
        Application.TaskPanes(wdTaskPaneFormatting).Visible = True
        
        'Selects three center boxes in Styles Pane Options, always
        ActiveDocument.FormattingShowFont = True
        ActiveDocument.FormattingShowParagraph = True
        ActiveDocument.FormattingShowNumbering = True
    
        'Shows all styles in document
        ActiveDocument.FormattingShowFilter = wdShowFilterStylesAvailable
        
        'Sorts styles alphabetically
        ActiveDocument.StyleSortMethod = wdStyleSortRecommended
        
        'Open Style Inspector
        Application.TaskPanes(wdTaskPaneStyleInspector).Visible = True
        
        ' Show nonprinting characters
        ActiveDocument.ActiveWindow.View.ShowAll = True
        
        ' Switch to Draft view with margin
        ActiveDocument.ActiveWindow.View.Type = wdNormalView
        ActiveDocument.ActiveWindow.StyleAreaWidth = InchesToPoints(1.5) 'Sets Styles margin area in draft view to 1.5 in
        
    Else    'we're already in draft view, switch back to print and close everything
        Application.TaskPanes(wdTaskPaneFormatting).Visible = False
        ' Close style inspector
        Application.TaskPanes(wdTaskPaneStyleInspector).Visible = False
        ' hide nonprinting characters
        ActiveDocument.ActiveWindow.View.ShowAll = False
        ' Switch to Normal (print) view
        ActiveDocument.ActiveWindow.View.Type = wdPrintView
        
    End If


End Sub
Private Sub Mac_BestStylesView()
    ' added by EW for v. 3.2
    ' Setup for Mac
    
    'if in normal/print view switch to draft and open other stuff
    If ActiveDocument.ActiveWindow.View.Type = wdNormalView Then
        ActiveDocument.ActiveWindow.View.Type = wdPrintView
        ActiveDocument.ActiveWindow.View.ShowAll = False
        Application.Dialogs(1755).Show ' its a toggle, so if it's open it will close with this
    Else
        With ActiveDocument.ActiveWindow
            .View.Type = wdNormalView
            .StyleAreaWidth = InchesToPoints(1.5)    'Sets Styles margin area in draft view to 1.5 in.
            .View.ShowAll = True
        End With
        Application.Dialogs(1755).Show
    End If

                           
End Sub
