Attribute VB_Name = "TagText"
Option Explicit

' By Erica Warren -- erica.warren@macmillan.com
' Tags all paragraphs with non-macmillan styles as TX or TX1
' And applies Space Before/After/Around styles for TX, FMTX, BMTX
' Though it doesn't detect space needed between the bottom of a table and the following 'graph

Sub TagText()
    ' Make sure we're always working with the right document
    Dim thisDoc As Document
    Set thisDoc = ActiveDocument
    
    ' Ask user if they want to tag space around extracts and such
    Dim blnTagSpaceAround As Boolean
    Dim strMessage As String
    
    strMessage = "Would you like to tag space around extracts/lists/etc.?" & vbNewLine & vbNewLine & _
        "If you're not sure, you probably don't need to do this."
        
    If MsgBox(strMessage, vbYesNo + vbQuestion + vbDefaultButton2, "Tag Space Around Extracts?") = vbNo Then
        blnTagSpaceAround = False
    Else
        blnTagSpaceAround = True
    End If
    
    ' ======== Start progress bar ========
    Dim sglPercentComplete As Single
    Dim strStatus As String
    Dim strTitle As String
    Dim objTagProgress As ProgressBar
    
    Set objTagProgress = New ProgressBar
    
    ' Add some fun statuses?
    
    strTitle = "Text Standard Tagging Macro"
    sglPercentComplete = 0.09
    strStatus = "* Getting started..." & vbNewLine
    
    Call UpdateBarAndWait(Bar:=objTagProgress, Status:=strStatus, Percent:=sglPercentComplete)
    
    
    ' ======= Run startup checks ========
    ' True means a check failed (e.g., doc protection on)
    If StartupSettings = True Then
        Call Cleanup
        Unload objTagProgress
        Exit Sub
    End If
    

    ' ======== Run Character styles macro =========
    ' If you tag paragraph styles before character styles, paragraphs w/ >50% direct formatting
    ' could have formatting stripped w/ no notification
    
    Call ActualCharStyles(oProgressChar:=objTagProgress, StartPercent:=0.09, TotalPercent:=0.35)
    
    ' ======== Start the tagging ========
    ' Rename built-in style that has parens
    thisDoc.Styles("Normal (Web)").NameLocal = "_"
    
    Dim lngParaCount As Long
    Dim a As Long
    Dim strCurrentStyle As String
    Dim strTX As String
    Dim strTX1 As String
    Dim strNewStyle As String
    Dim strParaStatus As String
    Dim sglStartingPercent As Single
    Dim sglTotalPercent As Single
    
    ' Making these variables so we don't get any input errors with the style names t/o
    strTX = "Text - Standard (tx)"
    strTX1 = "Text - Std No-Indent (tx1)"
    
    lngParaCount = thisDoc.Paragraphs.Count
    
    Dim myStyle As Style ' For error handlers
    On Error GoTo ErrorHandler1     ' adds this style if it is not in the document
    
    ' leave room in progress bar if we have to tag space around later
    If blnTagSpaceAround = True Then
        sglStartingPercent = 0.44   ' Percentage Progress Bar starts at for this loop
        sglTotalPercent = 0.3  ' Total percentage this loop will take in progress bar
    Else
        sglStartingPercent = 0.44
        sglTotalPercent = 0.47
    End If
    
    For a = 1 To lngParaCount
        
        If a Mod 100 = 0 Then
            ' Increment progress bar
            sglPercentComplete = (((a / lngParaCount) * sglTotalPercent) + sglStartingPercent)
            strParaStatus = "* Tagging non-Macmillan paragraphs with Text - Standard (tx): " & a & " of " & lngParaCount _
                & vbNewLine & strStatus
            Call UpdateBarAndWait(Bar:=objTagProgress, Status:=strParaStatus, Percent:=sglPercentComplete)
        End If
        
        strCurrentStyle = thisDoc.Paragraphs(a).Style
        'Debug.Print a & ": " & strCurrentStyle
        
        ' tag all non-Macmillan-style paragraphs with standard Macmillan styles
        ' Macmillan styles all end in close parens
        If Right(strCurrentStyle, 1) <> ")" Then
            ' If flush left, make No-Indent
            If thisDoc.Paragraphs(a).FirstLineIndent = 0 Then
                strNewStyle = strTX1
            Else
                strNewStyle = strTX
            End If
            
            ' Change the style of the paragraph in question
            ' This is where it will error if no style present
            thisDoc.Paragraphs(a).Style = strNewStyle
            
        End If
    Next a
    On Error GoTo 0
    
    strStatus = "* Tagging non-Macmillan paragraphs with Text - Standard (tx)..." & vbNewLine & strStatus
    
    ' Change Normal (Web) back
    thisDoc.Styles("Normal (Web),_").NameLocal = "Normal (Web)"
    
    ' ======== Check paras above and below each for space before/after styles ========
    ' but only if user selected that option!
    If blnTagSpaceAround = True Then
    
        On Error GoTo ErrorContinue ' Tests for error because style not present, continues with next style if so
        
        Dim strSearchStyle(1 To 6) As String
        Dim b As Long
        Dim c As Long
        Dim lngCount As Long
        Dim lngParaIndex As Long
        Dim strThisStyle As String
        Dim strPrevStyle As String
        Dim strNextStyle As String
        Dim strExtractStyle(1 To 9) As String
        Dim blnPrevStyle As Boolean
        Dim blnNextStyle As Boolean
        Dim strName As String
        Dim strCode As String
        Dim lngOpenParens As Long
        Dim lngCodeStart As Long
        Dim lngCodeLen As Long
        Dim lngNameLen As Long
        
        ' Styles that we're searching for to check for space before/after/around
        strSearchStyle(1) = "Text - Standard (tx)"
        strSearchStyle(2) = "Text - Std No-Indent (tx1)"
        strSearchStyle(3) = "FM Text (fmtx)"
        strSearchStyle(4) = "FM Text No-Indent (fmtx1)"
        strSearchStyle(5) = "BM Text (bmtx)"
        strSearchStyle(6) = "BM Text No-Indent (bmtx1)"
        
        ' Category of style that needs space around it - from style names
        strExtractStyle(1) = "Extract"
        strExtractStyle(2) = "Epigraph"
        strExtractStyle(3) = "List"
        strExtractStyle(4) = "Letter"
        strExtractStyle(5) = "Table"
        strExtractStyle(6) = "Sidebar"
        strExtractStyle(7) = "Box"
        strExtractStyle(8) = "Verse"
        strExtractStyle(9) = "Poem"
        
        sglStartingPercent = 0.74   ' Percentage Progress Bar starts at for this loop
        sglTotalPercent = 0.17  ' Total percentage this loop will take in progress bar
        
        For b = LBound(strSearchStyle()) To UBound(strSearchStyle())
            
            sglPercentComplete = (((b / UBound(strSearchStyle())) * sglTotalPercent) + sglStartingPercent)
            strStatus = "* Fixing space around " & strSearchStyle(b) & "..." & vbNewLine & strStatus
            Call UpdateBarAndWait(Bar:=objTagProgress, Status:=strStatus, Percent:=sglPercentComplete)
            

            
            Selection.HomeKey Unit:=wdStory     ' Have to start search from the beginning of the doc
            'Debug.Print "Searching for " & strSearchStyle(b) & " paragraphs"
            lngCount = 0
            
            With Selection.Find
                .ClearFormatting
                .Text = ""
                .Forward = True
                .Wrap = wdFindStop
                .Format = True
                .Style = thisDoc.Styles(strSearchStyle(b))
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                
                Do While .Execute(Forward:=True) = True And lngCount <= lngParaCount
                    lngCount = lngCount + 1     ' To prevent infinite loops
                    'Debug.Print lngCount
                    ' i.e. the overall index number of the current paragraph
                    lngParaIndex = thisDoc.Range(0, Selection.Paragraphs(1).Range.End).Paragraphs.Count
                    'Debug.Print "Current paragraph is " & lngParaIndex
    
                    strThisStyle = thisDoc.Paragraphs(lngParaIndex).Style
                    'Debug.Print "This style is: " & strThisStyle
    
                    ' Verify we're not looking at the first paragraph of the document, cuz can't get style of paragraph 0
                    If lngParaIndex > 1 Then
                        strPrevStyle = thisDoc.Paragraphs(lngParaIndex - 1).Style
                    Else    ' this is the first paragraph, just have it match current style so it won't change below
                        strPrevStyle = strThisStyle
                    End If
    
                    ' Verify we're not looking at the last paragraph of the document, cuz can't get style of para after it
                    If lngParaIndex < lngParaCount Then
                        strNextStyle = thisDoc.Paragraphs(lngParaIndex + 1).Style
                    Else    ' we're at the last paragraph and can't look after, so match current
                        strNextStyle = strThisStyle
                    End If
    
                    'Debug.Print "Previous style is: " & strPrevStyle
                    'Debug.Print "Next style is: " & strNextStyle
    '
    '                ' Only test if styles don't match; we know thisStyle is OK if they all match
                    If strThisStyle <> strPrevStyle Or strThisStyle <> strNextStyle Then
                        ' reset variables
                        blnPrevStyle = False
                        blnNextStyle = False
    
                        ' Test prev/next paras for extract styles
                        For c = LBound(strExtractStyle()) To UBound(strExtractStyle())
                            'Debug.Print "Searching for: " & strExtractStyle(c)
                            
                            ' Using InStr again so we don't have to list/loop through EVERY Extract/Epigraph/Etc style
                            If InStr(strPrevStyle, strExtractStyle(c)) > 0 Then
                                blnPrevStyle = True
                            End If
                            
                            ' Need two separate steps so we can capture instances where same style is before and after
                            If InStr(strNextStyle, strExtractStyle(c)) > 0 Then
                                blnNextStyle = True
                            End If
    
                            ' If both are true then stop looking
                            If blnPrevStyle = True And blnNextStyle = True Then
                                Exit For
                            End If
                        Next c
    
                        On Error GoTo 0
    
                        If blnPrevStyle = False And blnNextStyle = False Then   ' styles are fine as is, check next paragraph
                            GoTo ContinueLoop
                            'Debug.Print "Don't change style"
                        Else
                            ' pull out the current style name and code to use to create the new code
                            lngOpenParens = InStr(strThisStyle, "(")
                            lngCodeStart = lngOpenParens + 1
                            lngCodeLen = Len(strThisStyle) - lngCodeStart
                            strCode = Mid(strThisStyle, lngCodeStart, lngCodeLen)
                            lngNameLen = lngOpenParens - 1
                            strName = Mid(strThisStyle, 1, lngNameLen)
    
                            On Error GoTo ErrorNewStyle     ' If style doesn't exist, create it
    
                            ' create new style name based on extract paras before and/or after
                            If blnPrevStyle = True And blnNextStyle = False Then    ' need space before
                                strNewStyle = strName & "Space Before (#" & strCode & ")"
                            ElseIf blnPrevStyle = False And blnNextStyle = True Then    ' need space after
                                strNewStyle = strName & "Space After (" & strCode & "#)"
                            ElseIf blnPrevStyle = True And blnNextStyle = True Then     ' need space around
                                strNewStyle = strName & "Space Around (#" & strCode & "#)"
                            Else
                                strNewStyle = strThisStyle
                            End If
    
                            'Debug.Print "New style is: " & strNewStyle
    
                            ' change style of paragraph in question
                            thisDoc.Paragraphs(lngParaIndex).Style = strNewStyle
                            On Error GoTo 0
    
                        End If
                    End If
                    ' Now collapse the selection to the end of the paragraph, otherwise
                    ' the following Find just searches within the current selection and finds nothing or loops forever!
                    Selection.Collapse Direction:=wdCollapseEnd
ContinueLoop:
                Loop
            End With
ContinueNextB:
        Next b
    End If
    
    ' Cleanup stuff
    sglPercentComplete = 1#
    strStatus = "* Finishing up..." & vbNewLine & strStatus
    
    Call UpdateBarAndWait(Bar:=objTagProgress, Status:=strStatus, Percent:=sglPercentComplete)
    
    Call Cleanup
    
    Unload objTagProgress
    
    strMessage = "The Text Tagging Macro is complete!"
    MsgBox strMessage, vbOKOnly, "All done!"
    
    Exit Sub
    
ErrorHandler1:
    If Err.Number = 5834 Or Err.Number = 5941 Then  ' Style is not in doc
        Set myStyle = thisDoc.Styles.Add(Name:=strTX, Type:=wdStyleTypeParagraph)
        With myStyle
            '.QuickStyle = True ' not available for Mac
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            With .ParagraphFormat
                .LineSpacingRule = wdLineSpaceDouble
                .SpaceAfter = 0
                .SpaceBefore = 0
                
                With .Borders
                    Select Case strNewStyle
                        Case strTX
                            .OutsideLineStyle = wdLineStyleSingle
                            .OutsideLineWidth = wdLineWidth600pt
                        
                        Case strTX1
                            .OutsideLineStyle = wdLineStyleDouble
                            .OutsideLineWidth = wdLineWidth225pt
                            
                    End Select
                .OutsideColor = RGB(102, 204, 255)
                
                End With
            End With
        End With
        
        ' Now go back and try to assign that style again
        Resume
    
    Else
        Debug.Print "ErrorHandler1: " & Err.Number & " " & Err.Description
        On Error GoTo 0
        Call Cleanup
        Exit Sub
    End If
    
ErrorContinue:
    If Err.Number = 5834 Or Err.Number = 5941 Then  ' Style is not in doc
        GoTo ContinueNextB
    Else
        Debug.Print "ErrorContinue: " & Err.Number & " " & Err.Description
        On Error GoTo 0
        Call Cleanup
        Exit Sub
    End If
ErrorNewStyle:
    If Err.Number = 5834 Or Err.Number = 5941 Then  ' Style is not in doc
        ' So create style and apply formatting
        
        Set myStyle = thisDoc.Styles.Add(Name:=strNewStyle, Type:=wdStyleTypeParagraph)
        With myStyle
            .BaseStyle = strThisStyle
        
            If InStr(strNewStyle, "After") > 0 Then
                .ParagraphFormat.SpaceAfter = 18    ' in points
                If InStr(strNewStyle, "1") > 0 Then ' It's a no-indent style
                    .Borders.OutsideLineStyle = wdLineStyleDouble
                    .Borders.OutsideLineWidth = wdLineWidth150pt
                Else
                    .Borders.OutsideLineStyle = wdLineStyleSingle
                    .Borders.OutsideLineWidth = wdLineWidth300pt
                End If
            ElseIf InStr(strNewStyle, "Before") > 0 Then
                .ParagraphFormat.SpaceBefore = 18
                If InStr(strNewStyle, "1") > 0 Then ' It's a no-indent style
                    .Borders.OutsideLineStyle = wdLineStyleDouble
                    .Borders.OutsideLineWidth = wdLineWidth075pt
                Else
                    .Borders.OutsideLineStyle = wdLineStyleSingle
                    .Borders.OutsideLineWidth = wdLineWidth150pt
                End If
            ElseIf InStr(strNewStyle, "Around") > 0 Then
                .ParagraphFormat.SpaceBefore = 18
                .ParagraphFormat.SpaceAfter = 18
                If InStr(strNewStyle, "1") > 0 Then ' It's a no-indent style
                    .Borders.OutsideLineStyle = wdLineStyleDouble
                    .Borders.OutsideLineWidth = wdLineWidth300pt
                Else
                    .Borders.OutsideLineStyle = wdLineStyleSingle
                    .Borders.OutsideLineWidth = wdLineWidth450pt
                End If
            Else
                ' ?
            End If

        End With
        Resume
    Else
        Debug.Print "ErrorNewStyle: " & Err.Number & " " & Err.Description
        On Error GoTo 0
        Call Cleanup
        Exit Sub
    End If

End Sub

