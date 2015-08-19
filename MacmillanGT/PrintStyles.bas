Attribute VB_Name = "PrintStyles"
Option Explicit
Option Base 1
Sub PrintStyles()
    ' Prints current styles names to the left of each paragraph
    
    ' ===== DEPENDENCIES ==========================================================
    ' Before you run this, create a text box with the listed settings below, then select the
    ' text box and go to Insert > Text Box > Save Selection to Text Box Gallery (PC). In the
    ' Create New Building Block dialog that opens, name the Building Block "StyleNames1"
    ' and give it the category "Macmillan." Be sure to save it in the template at the path below.
    ' TEXT BOX SETTINGS:
    ' Layout > Position > Horizontal: Absolute position 0.13" to the right of Left Margin
    ' Layout > Position > Vertical: Absolute position 0" below paragraph
    ' Layout > Position > Lock Anchor = True
    ' Layout > Position > Allow overlap = true
    ' Layout > Text wrapping > Wrapping style = Square
    ' Layout > Text wrapping > Text wrap = right only
    ' Layout > Distance from text = 0" (each side)
    ' Layout > Size > Height: Absolute 0.4"
    ' Layout > Size > Width: Absolute 1.35
    
    ' ===== LIMITATIONS ==========================================================
    ' "Normal" can't be in-use as a style in the document
    ' If total margin size (left + right) is < 2 " paragraphs will reflow
    ' Doesn't work for endnotes/footnotes (can't add a drawing object to EN/FN)s
    
    ' ====== TO DO ======
    ' save original document (can use shared macros?)
    ' change template location to MacmillanGT.dotm
    ' add progress bar?

    
    Application.ScreenUpdating = False
    
    ' ===== Copy and Paste into a new doc ===================
    ' Paste below throws an alert that too many styles are being copied, so turn off alerts
    Dim lngOpt As Long
    lngOpt = Application.DisplayAlerts
    Application.DisplayAlerts = wdAlertsNone
    
    ' Copy the text of the document into a new document, so we don't screw up the original
    ActiveDocument.StoryRanges(wdMainTextStory).Select
    Selection.Copy
    
    ' Test if document has any text in it (1 = just a single paragraph return)
    Dim tempDoc As Document
    If Len(Selection) > 1 Then
        Set tempDoc = Documents.Add(Visible:=False) ' Can I set visibility to False here?
        tempDoc.Content.Paste
    Else
        MsgBox "Your document doesn't appear to have any content.", vbCritical, "Oops!"
        GoTo Cleanup
    End If
    
    ' ===== Set margins =================
    ' if possible, we want the total margin size to stay the same
    ' so that the paragraphs don't reflow
    ' note margins are in points, 72 pts = 1 inch
    Dim currentLeft As Long
    Dim currentRight As Long
    Dim currentTotal As Long
    
    With tempDoc.PageSetup
        currentLeft = .LeftMargin
        currentRight = .RightMargin
        currentTotal = currentLeft + currentRight
        .LeftMargin = 108   ' 1.5 inches
            If currentTotal >= 144 Then     ' 2 inches
                .RightMargin = currentTotal - 108   ' 1.5 inches
            Else
                .RightMargin = 36   '5 inches (minimum right margin)
            End If
    End With
    
    ' ==== Change Normal style formatting (since it will define the text boxes) =====
    ' But save settings first and then change back -- are these settings sticky?
    Dim currentSize As Long
    Dim currentName As String
    Dim currentSpace As Long
    
    With tempDoc.Styles("Normal")
        currentSize = .Font.Size
        currentName = .Font.Name
        currentSpace = .ParagraphFormat.SpaceAfter
        .Font.Size = 7
        .Font.Name = "Calibri"
        .ParagraphFormat.SpaceAfter = 0
    End With
    
    Dim strPath As String
    Dim objTemplate As Template
    Dim objBB As BuildingBlock
    Dim strStyle As String
    Dim lngTextBoxes As Long
    Dim a As Long
    Dim b As Long
    
    ' This is the template where the building block is saved
    strPath = Environ("APPDATA") & "\Microsoft\Word\STARTUP\PrintStyles.dotm"
    Set objTemplate = Templates(strPath)
    
    ' Access the building block through the type and category
    ' NOTE the text box building block has to already be created in the template.
    Set objBB = objTemplate.BuildingBlockTypes(wdTypeTextBox).Categories("Macmillan").BuildingBlocks("StyleNames1")
    
    ' Count the number of current text boxes, because the index number of the new ones
    ' will be offset by that amount
    lngTextBoxes = tempDoc.Shapes.Count
    
    For a = 1 To tempDoc.Paragraphs.Count
    
        tempDoc.Paragraphs(a).Range.Select
        strStyle = Selection.Style
        Selection.Collapse Direction:=wdCollapseStart
        objBB.Insert Where:=Selection.Range
        tempDoc.Shapes(a + lngTextBoxes).TextFrame.TextRange.Text = strStyle
    
    Next a

    ' Now open the print dialog so user can print the document.
    Dialogs(wdDialogFilePrint).Show
    
    ' reset Normal style because I'm not sure if it's sticky or not
    With tempDoc.Styles("Normal")
        .Font.Size = currentSize
        .Font.Name = currentName
        .ParagraphFormat.SpaceAfter = currentSpace
    End With
    
    ' Close newly created doc w/o saving
    tempDoc.Close wdDoNotSaveChanges
        
Cleanup:
    With Application
        .DisplayAlerts = lngOpt
        .ScreenUpdating = True
        .ScreenRefresh
    End With
End Sub


