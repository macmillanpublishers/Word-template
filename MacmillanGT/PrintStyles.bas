Attribute VB_Name = "PrintStyles"
Option Explicit
Option Base 1
Sub PrintStyles()
    ' Created by Erica Warren -- erica.warren@macmillan.com
    
    ' places current styles names to the left of each paragraph in the margin and prints document
    
    ' ===== DEPENDENCIES ==========================================================
    ' Requires modules AttachTemplateMacro (and thus also Macmillan style templates) and SharedMacros
    ' Before you run this, create a text box with the listed settings below, then select the
    ' text box and go to Insert > Text Box > Save Selection to Text Box Gallery (Word 2013). In the
    ' Create New Building Block dialog that opens, name the Building Block "StyleNames1" and
    ' give it the category "Macmillan." Be sure to save it in the template at the path below.
    
    ' TEXT BOX SETTINGS:
    ' Layout > Position > Horizontal: Absolute position 0.13" to the right of Left Margin
    ' Layout > Position > Vertical: Absolute position 0" below paragraph
    ' Layout > Position > Lock Anchor = True
    ' Layout > Position > Allow overlap = true
    ' Layout > Text wrapping > Wrapping style = Square
    ' Layout > Text wrapping > Text wrap = right only
    ' Layout > Distance from text = 0" (each side)
    ' Layout > Size > Height: Absolute 0.4"
    ' Layout > Size > Width: Absolute 1.35"
    ' Format > Shape Styles > Shape Outline: No outline
    
    ' ===== LIMITATIONS ==========================================================
    ' "Normal" can't be in-use as a style in the document
    ' If total margin size (left + right) is < 2 " paragraphs will reflow
    ' Doesn't work for endnotes/footnotes (can't add a drawing object to EN/FNs)
    
    '=================================================
    '                  Timer Start                  '|
    Dim StartTime As Double                         '|
    Dim SecondsElapsed As Double                    '|
                                                    '|
    'Remember time when macro starts                '|
    StartTime = Timer                               '|
    '=================================================
    
    ' ====== Check if doc is saved/protected ================
    If CheckSave = True Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' ===== Progress Bar / Status Bar ========================
    Dim currentStatusBar As Boolean
    currentStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    
    Dim sglPercentComplete As Single
    Dim strStatus As String
    Dim strTitle As String
       
    ' Enter fun array here
    
    strTitle = "Print Styles in Margin"
    sglPercentComplete = 0.01
    strStatus = "* Getting started..."
    
    #If Mac Then
        Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
        DoEvents
    #Else
        Dim objProgressPrint As ProgressBar
        Set objProgressPrint = New ProgressBar
        objProgressPrint.Title = strTitle
        objProgressPrint.Show
        objProgressPrint.Increment sglPercentComplete, strStatus
        Doze 50
    #End If
    
    ' ===== Copy and Paste into a new doc ===================
    ' Paste below throws an alert that too many styles are being copied, so turn off alerts
    ' Also make sure paste settings maintain styles
    Dim lngOpt As Long
    Dim lngPasteStyled As Long
    Dim lngPasteFormat As Long
    
    With Application
    ' record current settings to reset in Cleanup
        lngOpt = .DisplayAlerts
        'lngPasteStyled = .Options.PasteFormatBetweenStyledDocuments ' not available on Mac
        'lngPasteFormat = .Options.PasteFormatBetweenDocuments  ' not available on Mac
        .DisplayAlerts = wdAlertsMessageBox
        '.Options.PasteFormatBetweenStyledDocuments = wdKeepSourceFormatting
        '.Options.PasteFormatBetweenDocuments = wdKeepSourceFormatting
    End With
    
    ' ===== Create new version of this document to manipulate ============
    sglPercentComplete = 0.03
    strStatus = "* Creating dupe document to tag with style names..." & vbNewLine & strStatus
    
    #If Mac Then
        Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
        DoEvents
    #Else
        objProgressPrint.Title = strTitle
        objProgressPrint.Show
        objProgressPrint.Increment sglPercentComplete, strStatus
        Doze 50
    #End If
    
    ' Copy the text of the document into a new document, so we don't screw up the original
    ' Needs to have the BoundMS template attached before copying so the styles match
    ' the new document later, or won't copy any styles
    Dim currentTemplate As String
    Dim currentDoc As Document
    Set currentDoc = ActiveDocument
    ' Record current template
    currentTemplate = currentDoc.AttachedTemplate
    
    ' Attach BoundMS template to original doc, then copy contents
    Call AttachTemplateMacro.zz_AttachBoundMSTemplate
    currentDoc.StoryRanges(wdMainTextStory).Copy
    
    Dim tempDoc As Document
    ' Create a new document
    Set tempDoc = Documents.Add '(Visible:=False) ' Can I set visibility to False here on Mac?
    ' Add Macmillan styles with no color guides (because if we don't add them,
    ' we get an error that there are too many styles to paste and it just pastes
    ' all with Normal style)
    tempDoc.Activate
    Call AttachTemplateMacro.zz_AttachBoundMSTemplate
    tempDoc.Content.PasteAndFormat wdFormatOriginalFormatting
        
    ' ===== Set margins =================
    sglPercentComplete = 0.05
    strStatus = "* Adjusting margins to fit style names..." & vbNewLine & strStatus
    
    #If Mac Then
        Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
        DoEvents
    #Else
        objProgressPrint.Title = strTitle
        objProgressPrint.Show
        objProgressPrint.Increment sglPercentComplete, strStatus
        Doze 50
    #End If
    
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
                .RightMargin = 36   '.5 inches (minimum right margin)
            End If
    End With
    
    ' ==== Change Normal style formatting (since it will define the text boxes) =====
    sglPercentComplete = 0.07
    strStatus = "* Setting format for style names..." & vbNewLine & strStatus
    
    #If Mac Then
        Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
        DoEvents
    #Else
        objProgressPrint.Title = strTitle
        objProgressPrint.Show
        objProgressPrint.Increment sglPercentComplete, strStatus
        Doze 50
    #End If
        
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
    
    ' ==== Add style names in margin ==================================
    Dim strPath As String
    Dim objTemplate As Template
    Dim objBB As BuildingBlock
    Dim strStyle As String
    Dim lngTextBoxes As Long
    Dim activeParas As Long
    Dim a As Long
    Dim b As Long
    
    ' This is the template where the building block is saved
    'strPath = Environ("APPDATA") & "\Microsoft\Word\STARTUP\MacmillanGT.dotm"
    'If IsItThere(strPath) = True Then
    '    Set objTemplate = Templates(strPath)
    'Else
    '    MsgBox "I can't find the Macmillan template, sorry."
    '    GoTo Cleanup
   ' End If
    
    ' Access the building block through the type and category
    ' NOTE the text box building block has to already be created in the template.
    ' Set objBB = objTemplate.BuildingBlockTypes(wdTypeTextBox).Categories("Macmillan").BuildingBlocks("StyleNames1") ' building blocks don't exist in Word 2011 Mac
    
    ' Count the number of current text boxes etc., because the index number of the new ones
    ' will be offset by that amount
    Dim strStatusLoop As String
    lngTextBoxes = tempDoc.Shapes.Count
    activeParas = tempDoc.Paragraphs.Count
    
    For a = 1 To activeParas
        If a Mod 50 = 0 Then
            sglPercentComplete = Round((((a / activeParas) * 0.85) + 0.1), 2)
            strStatusLoop = "* Adding style names to paragraph " & a & " of " & activeParas & "..." & vbNewLine & strStatus
            
            #If Mac Then
                Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatusLoop
                DoEvents
            #Else
                objProgressPrint.Title = strTitle
                objProgressPrint.Show
                objProgressPrint.Increment sglPercentComplete, strStatusLoop
                Doze 50
            #End If
        End If
    
        tempDoc.Paragraphs(a).Range.Select
        strStyle = Selection.Style
        Selection.Collapse Direction:=wdCollapseStart
        'objBB.Insert Where:=Selection.Range
        Dim newBox As Shape
        Set newBox = tempDoc.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Left:=InchesToPoints(0.13), Top:=0, Height:=InchesToPoints(0.4), _
            Width:=InchesToPoints(1.35))
        With newBox
            .RelativeHorizontalPosition = wdRelativeHorizontalPositionLeftMarginArea
            .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
            .Top = 0
            .Line.Visible = False
            .LockAnchor = True
            .TextFrame.TextRange.Text = strStyle
            
            With .WrapFormat
                .Type = wdWrapSquare
                .Side = wdWrapBoth
                .DistanceTop = 0
                .DistanceBottom = 0
                .DistanceLeft = 0
                .DistanceRight = 0
            End With
            
        End With
        
    Next a
    
    strStatus = "* Adding style names to margin..." & vbNewLine & strStatus

    ' Now open the print dialog so user can print the document.
    sglPercentComplete = 0.97
    strStatus = strStatus & "* Printing document with style names in  margin..." & vbNewLine
    
    #If Mac Then
        Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
        DoEvents
    #Else
        objProgressPrint.Title = strTitle
        objProgressPrint.Show
        objProgressPrint.Increment sglPercentComplete, strStatus
        Doze 50
    #End If
    
    Dialogs(wdDialogFilePrint).Show
    
    ' Cleanup
    sglPercentComplete = 1
    strStatus = strStatus & "* Finishing up..." & vbNewLine
    
    #If Mac Then
        Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
        DoEvents
    #Else
        objProgressPrint.Title = strTitle
        objProgressPrint.Show
        objProgressPrint.Increment sglPercentComplete, strStatus
        Doze 50
    #End If
    
Cleanup:
    ' Close newly created doc w/o saving
    Dim tempDocPath As String
    tempDocPath = tempDoc.Path

    If IsItThere(tempDocPath) = True Then
        ' reset Normal style because I'm not sure if it's sticky or not
        With tempDoc.Styles("Normal")
            .Font.Size = currentSize
            .Font.Name = currentName
            .ParagraphFormat.SpaceAfter = currentSpace
        End With
        ' Close temo doc without saving
        tempDoc.Close wdDoNotSaveChanges
    End If
    
    ' Return original document to original template
    currentDoc.Activate
    Call AttachTemplateMacro.AttachMe(TemplateName:=currentTemplate)
    
    #If Mac Then
        'Nothing
    #Else
        Unload objProgressPrint
    #End If
    
    ' Reset settings to original
    With Application
        .DisplayAlerts = lngOpt
        '.Options.PasteFormatBetweenStyledDocuments = lngPasteStyled
        '.Options.PasteFormatBetweenDocuments = lngPasteFormat
        .DisplayStatusBar = currentStatusBar
        .ScreenUpdating = True
        .ScreenRefresh
    End With
    
    '============================================================================
    '----------------------Timer End-------------------------------------------
    ''''Determine how many seconds code took to run
      SecondsElapsed = Round(Timer - StartTime, 2)
    
    ''''Notify user in seconds
      Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"
    '============================================================================


End Sub


