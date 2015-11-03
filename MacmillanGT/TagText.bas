Attribute VB_Name = "TagText"
Option Explicit

' By Erica Warren -- erica.warren@macmillan.com
' Tags all paragraphs with non-macmillan styles as TX or TX1
' And applies Space Before/After/Around styles for TX, FMTX, BMTX

Sub TagText()
    ' Make sure we're always working with the right document
    Dim thisDoc As Document
    Set thisDoc = ActiveDocument
    
    ' Add check if doc is saved
    If CheckSave = True Then
        Exit Sub
    End If
    
    ' ======== Start the tagging ========
    ' Rename built-in style with parens
    thisDoc.Styles("Normal (Web)").NameLocal = "_"
    
    Dim lngParaCount As Long
    Dim a As Long
    Dim strCurrentStyle As String
    lngParaCount = thisDoc.Paragraphs.Count
    
    ' ======== Tag non-macmillan styles with TX or TX1 ========
    On Error GoTo ErrorHandler1
    For a = 1 To lngParaCount
        strCurrentStyle = thisDoc.Paragraphs(a).Style
        
        ' Macmillan styles all end in close parens
        If Right(strCurrentStyle, 1) <> ")" Then
            ' If flush left, make No-Indent
            If thisDoc.Paragraphs(a).FirstLineIndent = 0 Then
                thisDoc.Paragraphs(a).Style = "Text - Std No-Indent (tx1)"
            Else
                thisDoc.Paragraphs(a).Style = "Text - Standard (tx)"
            End If
        End If
    Next a
    On Error GoTo 0
    
    ' Change Normal (Web) back
    thisDoc.Styles("Normal (Web),_").NameLocal = "Normal (Web)"
    
    ' ======== Check paras above and below each for space before/after styles ========
    ' the styles we'll be searching through for space before/after needed
    Dim strSearchStyle(1 To 6) As String
    Dim b As Long
    Dim c As Long
    Dim lngCount As Long
    Dim lngParaIndex As Long
    Dim strThisStyle As String
    Dim strPrevStyle As String
    Dim strNextStyle As String
    Dim strExtractStyle(1 To 9) As String
    Dim blnPrevExt As Boolean
    Dim blnNextExt As Boolean
    Dim strNewStyle As String
    Dim strName As String
    Dim strCode As String
    Dim lngOpenParens As Long
    Dim lngCodeStart As Long
    Dim lngCodeLen As Long
    Dim lngNameLen As Long
    
    strSearchStyle(1) = "Text - Standard (tx)"
    strSearchStyle(2) = "Text - Std No-Indent (tx1)"
    strSearchStyle(3) = "FM Text (fmtx)"
    strSearchStyle(4) = "FM Text No-Indent (fmtx1)"
    strSearchStyle(5) = "BM Text (bmtx)"
    strSearchStyle(6) = "BM Text (bmtx1)"
    
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
    
    ' Does this need to be Selection.Find? If so, use the following first
    ' Selection.HomeKey Unit:=wdStory
    On Error GoTo ErrorContinue
    For b = LBound(strSearchStyle()) To UBound(strSearchStyle())
        lngCount = 0
        With thisDoc.Range.Find
            .ClearFormatting
            .Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .Style = strSearchStyle(b)  'OR .Style = thisDoc.Styles(strSearchStyle(b)) ?
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            Do While .Execute(Forward:=True) = True And lngCount < 10000
                lngCount = lngCount + 1     ' To prevent infinite loops
                ' This below might have to be Selection ... will that mess up the Range.Find above?
                lngParaIndex = thisDoc.Range(0, thisDoc.Range.Paragraphs(1).Range.End).Paragraphs.Count
                Debug.Print "Current paragraph is " & lngParaIndex
                
                strThisStyle = thisDoc.Paragraphs(lngParaIndex).Style
                
                ' Verify we're not looking at the first paragraph of the document
                If lngParaIndex > 1 Then
                    strPrevStyle = thisDoc.Paragraphs(lngParaIndex - 1).Style
                Else    ' this is the first paragraph, just have it match current style so it won't change below
                    strPrevStyle = strThisStyle
                End If
                
                ' Verify we're not looking at the last paragraph of the document
                If lngParaIndex < lngParaCount Then
                    strNextStyle = thisDoc.Paragraphs(lngParaIndex + 1).Style
                Else    ' we're at the last paragraph anda can't look after, so match current
                    strNextStyle = strThisStyle
                End If
                
                ' Only test if styles don't match; we know thisStyle is OK
                If strThisStyle <> strPrevStyle Or strThisStyle <> strNextStyle Then
                    ' reset variables
                    blnPrevStyle = False
                    blnNextStyle = False
                    
                    ' Test prev/next paras for extract styles
                    For c = LBound(strExtractStyle()) To UBound(strExtractStyle())
                        If InStr(strPrevStyle, strExtractStyle(c)) > 0 Then
                            blnPrevStyle = True
                        ElseIf InStr(strNextStyle, strExtractStyle(c)) > 0 Then
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
                    Else
                        ' pull out the style name and code to use to create the new code
                        lngOpenParens = InStr(strThisStyle, "(")
                        lngCodeStart = lngOpenParens + 1
                        lngCodeLen = Len(strThisStyle) - lngCodeStart
                        strCode = Mid(strThisStyle, lngCodeStart, lngCodeLen)
                        lngNameLen = lngOpenParens - 1
                        strName = Mid(strThisStyle, 1, lngNameLen)
                        
                        On Error GoTo ErrorNewStyle
                        
                        ' create new style name based on extract paras before and/or after
                        If blnPrevStyle = True And blnNextStyle = False Then    ' need space after
                            strNewStyle = strName & "Space After (" & strCode & "#)"
                        ElseIf blnPrevStyle = False And blnNextStyle = True Then    ' need space before
                            strNewStyle = strName & "Space Before (#" & strCode & ")"
                        ElseIf blnPrevStyle = True And blnNextStyle = True Then     ' need space around
                            strNewStyle = strName & "Space Around (#" & strCode & "#)"
                        Else
                            strNewStyle = strThisStyle
                        End If
                        
                        ' change style of paragraph in question
                        thisDoc.Paragraphs(lngParaIndex).Style = strNewStyle
                        On Error GoTo 0
                        
                    End If
                End If
ContinueLoop:
            Loop
        End With
ContinueNextB:
    Next b
    
    Exit Sub
    
ErrorHandler1:
    If Err.Number = 5834 Or Err.Number = 5941 Then  ' Style is not in doc
        MsgBox "This macro requires the Macmillan styles in the document. Please add the Macmillan styles and try again."
        Exit Sub
    Else
        Debug.Print "ErrorHandler1: " & Err.Number & " " & Err.Description
        Exit Sub
    End If
    
ErrorContinue:
    If Err.Number = 5834 Or Err.Number = 5941 Then  ' Style is not in doc
        GoTo ContinueNextB
    Else
        Debug.Print "ErrorContinue: " & Err.Number & " " & Err.Description
        Exit Sub
    End If
ErrorNewStyle:
    If Err.Number = 5834 Or Err.Number = 5941 Then  ' Style is not in doc
        ' So create style and apply formatting
        Dim myStyle As Style
        
        Set myStyle = ActiveDocument.Styles.Add(Name:=strNewStyle, Type:=wdStyleTypeParagraph)
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
        Exit Sub
    End If

End Sub
