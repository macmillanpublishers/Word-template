Attribute VB_Name = "torCastoffMacro"
Option Explicit

Sub torCastoff()

Dim charCount As Long
Dim designChar As Integer
Dim pageCount As Integer

'Get character count with space from Word doc, divide by avg. char. count of design to get page count
charCount = ActiveDocument.Range.ComputeStatistics(wdStatisticCharactersWithSpaces)
designChar = 1275               'average character count with spaces per page of print text design
pageCount = charCount / designChar

'Debug.Print pageCount

'search for page breaks, add a page for each
With ActiveDocument.Range.Find
    .ClearFormatting
    .Text = "^m"
    .Replacement.Text = "^m"
    .Forward = True
    .Wrap = wdFindStop
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
Do While .Execute(Forward:=True) = True
    pageCount = pageCount + 1
Loop
End With

'Debug.Print pageCount

'search for space breaks
Dim sbCount As Long
sbCount = 0

With ActiveDocument.Range.Find
    .ClearFormatting
    .Text = "^p^p"
    .Replacement.Text = "^p^p"
    .Forward = True
    .Wrap = wdFindStop
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
Do While .Execute(Forward:=True) = True And sbCount < 500   'to prevent infinite loops
    sbCount = sbCount + 1
Loop
End With

'Debug.Print sbCount

'divide sbCount by 26 (26 lines per page in print design)
sbCount = sbCount / 26

'Debug.Print sbCount

pageCount = pageCount + sbCount

'Debug.Print pageCount

'Test if pageCount is odd or even
If ((pageCount Mod 2) = 0) Then ' Even = True
    pageCount = pageCount
Else                            ' pageCount is odd, must round up 1 page.
    pageCount = pageCount + 1
End If
    
'Debug.Print pageCount

If pageCount > 56 Then
    MsgBox "Your book will be approximately " & pageCount & " pages using the Tor.com automated conversion tool."
Else
    MsgBox "Your book will be approximately " & pageCount & " pages using the Tor.com automated conversion tool." & vbNewLine & _
    vbNewLine & "Note that books 48 pages and shorter will be saddle stitched (no spine)."
End If

End Sub





