Attribute VB_Name = "Reports"
Option Explicit
Option Base 1
Sub BookmakerReqs()
'-----------------------------------------------------------

'Created by Erica Warren - erica.warren@macmillan.com
'4/21/2015: breaking down into individual subs and functions for each step
'4/10/2015: adding handling of track changes
'4/3/2015: adding Imprint Line requirement
'3/27/2015: converts solo CNs to CTs
'           page numbers added to Illustrations List
'           Added style report WITH character styles
'3/20/2015: Added check if template is attached
'3/17/2015: Added Illustrations List
'3/16/2015: Fixed error creating text file, added title/author/isbn confirmation

'------------------------------------------------------------

'------------------Time Start-----------------
Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
StartTime = Timer

Application.ScreenUpdating = False

'-----------run preliminary error checks------------
Dim exitOnError As Boolean
exitOnError = srErrorCheck()

If exitOnError <> False Then
Exit Sub
End If

'-------Delete content controls on PC------------------------
'Has to be a separate sub because these objects don't exist in Word 2011 Mac and it won't compile
Dim TheOS As String
TheOS = System.OperatingSystem

If Not TheOS Like "*Mac*" Then
    Call DeleteContentControlPC
End If

'-------Deal with Track Changes and Comments----------------
If FixTrackChanges = False Then
    Application.ScreenUpdating = True
    Exit Sub
End If


'-------Count number of occurences of each required style----
Dim styleCount() As Variant

styleCount = CountReqdStyles()

If styleCount(1) = 100 Then     'Then count got stuck in a loop, gave message to user in last function
    Application.ScreenUpdating = True
    Exit Sub
End If
            
'------------Convert solo Chap Number paras to Chap Title-------
If styleCount(4) > 0 And styleCount(5) = 0 Then         'If Chap Num > 0 and Chap Title = 0
    Call FixSoloCNs
End If

'--------Get title/author/isbn/imprint text from document-----------
Dim strMetadata As String
strMetadata = GetMetadata

            
'-------------------Get Illustrations List from Document-----------
Dim strIllustrationsList As String
strIllustrationsList = IllustrationsList


'-------------------Get list of good and bad styles from document---------
Dim arrGoodBadStyles() As Variant
Dim strGoodStylesList As String
Dim strBadStylesList As String

'returns array with 2 elements, 1: good styles list, 2: bad styles list
arrGoodBadStyles = GoodBadStyles(torDOTcom:=True)

strGoodStylesList = arrGoodBadStyles(1)
strBadStylesList = arrGoodBadStyles(2)


'-------------------Create error report----------------------------
Dim strErrorList As String
strErrorList = CreateErrorList(strBadStylesList, styleCount)

'------Create Report File-------------------------------
Dim strSuffix As String
strSuffix = "BookmakerReport" ' suffix for the report file
Call CreateReport(strErrorList, strMetadata, strIllustrationsList, strGoodStylesList, strSuffix)

Application.ScreenUpdating = True

'----------------------Timer End-------------------------------------------
''''Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

''''Notify user in seconds
  Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"

End Sub
Sub MacmillanStyleReport()

Application.ScreenUpdating = False

'-----------Turn off track changes--------
Dim currentTracking As Boolean
currentTracking = ActiveDocument.TrackRevisions
ActiveDocument.TrackRevisions = False

'-----------run preliminary error checks------------
Dim exitOnError As Boolean
exitOnError = srErrorCheck()

If exitOnError <> False Then
Exit Sub
End If

'-------Delete content controls on PC------------------------
'Has to be a separate sub because these objects don't exist in Word 2011 Mac and it won't compile
Dim TheOS As String
TheOS = System.OperatingSystem

If Not TheOS Like "*Mac*" Then
    Call DeleteContentControlPC
End If

'-------Count number of occurences of each required style----
Dim styleCount() As Variant

styleCount = CountReqdStyles()

If styleCount(1) = 100 Then     'Then count got stuck in a loop, gave message to user in last function
    Application.ScreenUpdating = True
    Exit Sub
End If
            
'------------Convert solo Chap Number paras to Chap Title-------
If styleCount(4) > 0 And styleCount(5) = 0 Then         'If Chap Num > 0 and Chap Title = 0
    Call FixSoloCNs
End If

'--------Get title/author/isbn/imprint text from document-----------
Dim strMetadata As String
strMetadata = GetMetadata

            
'-------------------Get Illustrations List from Document-----------
Dim strIllustrationsList As String
strIllustrationsList = IllustrationsList


'-------------------Get list of good and bad styles from document---------
Dim arrGoodBadStyles() As Variant
Dim strGoodStylesList As String
Dim strBadStylesList As String

'returns array with 2 elements, 1: good styles list, 2: bad styles list
arrGoodBadStyles = GoodBadStyles(torDOTcom:=False)

strGoodStylesList = arrGoodBadStyles(1)
strBadStylesList = arrGoodBadStyles(2)

'-------------------Create error report----------------------------
Dim strErrorList As String
strErrorList = CreateErrorList(strBadStylesList, styleCount)


'-----------------------create text file------------------------------
Dim strSuffix As String
strSuffix = "StyleReport"       'suffix for report file, no spaces
Call CreateReport(strErrorList, strMetadata, strIllustrationsList, strGoodStylesList, strSuffix)

ActiveDocument.TrackRevisions = currentTracking         'Return track changes to the original setting


End Sub
Private Function GoodBadStyles(torDOTcom As Boolean) As Variant
'Creates a list of Macmillan styles in use
'And a separate list of non-Macmillan styles in use

Dim activeDoc As Document
Set activeDoc = ActiveDocument
Dim stylesGood() As String
Dim stylesGoodLong As Long
stylesGoodLong = 400                                    'could maybe reduce this number
ReDim stylesGood(stylesGoodLong)
Dim stylesBad() As String
ReDim stylesBad(1 To 100) 'could maybe reduce this number too
Dim styleGoodCount As Integer
Dim styleBadCount As Integer
Dim styleBadOverflow As Boolean
Dim activeParaCount As Integer
Dim J As Integer, K As Integer, L As Integer
Dim paraStyle As String
'''''''''''''''''''''
Dim activeParaRange As Range
Dim pageNumber As Integer


'Alter built-in Normal (Web) style temporarily (later, maybe forever?)
ActiveDocument.Styles("Normal (Web)").NameLocal = "_"

' Collect all styles being used
styleGoodCount = 0
styleBadCount = 0
styleBadOverflow = False
activeParaCount = activeDoc.paragraphs.Count
For J = 1 To activeParaCount
    'Next two lines are for the status bar
    Application.StatusBar = "Checking paragraph: " & J & " of " & activeParaCount
    If J Mod 100 = 0 Then DoEvents
    
    paraStyle = activeDoc.paragraphs(J).Style
        'If InStrRev(paraStyle, ")", -1, vbTextCompare) Then        'ALT calculation to "Right", can speed test
    If Right(paraStyle, 1) = ")" Then
        For K = 1 To styleGoodCount
            If paraStyle = stylesGood(K) Then                   'stylereport bug fix #1  v. 3.1
                K = styleGoodCount                              'stylereport bug fix #1    v. 3.1
                Exit For                                        'stylereport bug fix #1   v. 3.1
            End If                                              'stylereport bug fix #1   v. 3.1
        Next K
        If K = styleGoodCount + 1 Then
            styleGoodCount = K
            stylesGood(styleGoodCount) = paraStyle
        End If
    Else
        For L = 1 To styleBadCount
            'If paraStyle = stylesBad(L) Then Exit For                  'Not needed, since we want EVERY instance of bad style
        Next L
        If L > 100 Then
                styleBadOverflow = True
            Exit For
        End If
        If L = styleBadCount + 1 Then
            styleBadCount = L
            Set activeParaRange = ActiveDocument.paragraphs(J).Range
            pageNumber = activeParaRange.Information(wdActiveEndPageNumber)                 'alt: (wdActiveEndAdjustedPageNumber)
            stylesBad(styleBadCount) = "**ERROR: Non-Macmillan style on page " & pageNumber & " (Paragraph " & J & "): " _
                & vbTab & paraStyle & vbNewLine & vbNewLine
        End If
    End If
Next J
 
'Change Normal (Web) back (if you want to)
ActiveDocument.Styles("Normal (Web),_").NameLocal = "Normal (Web)"

'Sort good styles
If K <> 0 Then
ReDim Preserve stylesGood(1 To styleGoodCount)
WordBasic.SortArray stylesGood()
End If

'Create single string for good styles
Dim strGoodStyles As String
For K = LBound(stylesGood()) To UBound(stylesGood())
    strGoodStyles = strGoodStyles & stylesGood(K) & vbNewLine
Next K

'Debug.Print strGoodStyles

'Create single string for bad styles
Dim strBadStyles As String
ReDim Preserve stylesBad(1 To styleBadCount)
For L = LBound(stylesBad()) To UBound(stylesBad())
    strBadStyles = strBadStyles & stylesBad(L)
Next L

'Debug.Print strBadStyles

'-------------------get list of good character styles--------------

Dim charStyles As String
Dim styleNameM(1 To 21) As String        'declare number in array
Dim M As Integer

styleNameM(1) = "span italic characters (ital)"
styleNameM(2) = "span boldface characters (bf)"
styleNameM(3) = "span small caps characters (sc)"
styleNameM(4) = "span underscore characters (us)"
styleNameM(5) = "span superscript characters (sup)"
styleNameM(6) = "span subscript characters (sub)"
styleNameM(7) = "span bold ital (bem)"
styleNameM(8) = "span smcap ital (scital)"
styleNameM(9) = "span smcap bold (scbold)"
styleNameM(10) = "span symbols (sym)"
styleNameM(11) = "span accent characters (acc)"
styleNameM(12) = "span cross-reference (xref)"
styleNameM(13) = "span hyperlink (url)"
styleNameM(14) = "span material to come (tk)"
styleNameM(15) = "span carry query (cq)"
styleNameM(16) = "span preserve characters (pre)"
styleNameM(17) = "bookmaker force page break (br)"
styleNameM(18) = "bookmaker keep together (kt)"
styleNameM(19) = "span ISBN (isbn)"
styleNameM(20) = "span symbols ital (symi)"
styleNameM(21) = "span symbols bold (symb)"

'Move selection back to start of document
Selection.HomeKey Unit:=wdStory

For M = 1 To UBound(styleNameM())
    With Selection.Find
        .Style = ActiveDocument.Styles(styleNameM(M))
        .Wrap = wdFindContinue
        .Format = True
    End With
    If Selection.Find.Execute = True Then
        charStyles = charStyles & styleNameM(M) & vbNewLine
    End If
Next M

'Move selection back to start of document
Selection.HomeKey Unit:=wdStory

'Add character styles to Good styles list
strGoodStyles = strGoodStyles & charStyles

'If this is for the Tor.com Bookmaker toolchain, test if only those styles used
Dim strTorBadStyles As String
If torDOTcom = True Then
    strTorBadStyles = BadTorStyles
    strBadStyles = strBadStyles & strTorBadStyles
End If

'Debug.Print strGoodStyles
'Debug.Print strBadStyles

'Add both good and bad styles lists to an array to pass back to original sub
Dim arrFinalLists() As Variant
ReDim arrFinalLists(1 To 2)

arrFinalLists(1) = strGoodStyles
arrFinalLists(2) = strBadStyles

GoodBadStyles = arrFinalLists

End Function
Private Function srErrorCheck()

srErrorCheck = False
Dim mainDoc As Document
Set mainDoc = ActiveDocument
Dim iReply As Integer

'-------Check if Macmillan template is attached--------------
Dim currentTemplate As String
Dim ourTemplate1 As String
Dim ourTemplate2 As String
Dim ourTemplate3 As String

currentTemplate = ActiveDocument.BuiltInDocumentProperties(wdPropertyTemplate)
ourTemplate1 = "macmillan.dotm"
ourTemplate2 = "macmillan_NoColor.dotm"
ourTemplate3 = "MacmillanCoverCopy.dotm"

'Debug.Print "Current template is " & currentTemplate & vbNewLine

If currentTemplate <> ourTemplate1 Then
    If currentTemplate <> ourTemplate2 Then
        If currentTemplate <> ourTemplate3 Then
            MsgBox "Please attach the Macmillan Style Template to this document and run the macro again."
            Exit Function
        End If
    End If
End If

'-----make sure document is saved
Dim docSaved As Boolean
docSaved = mainDoc.Saved
If docSaved = False Then
    iReply = MsgBox("Your document '" & mainDoc & "' contains unsaved changes." & vbNewLine & vbNewLine & _
        "Click OK, and I will save the document and run the report." & vbNewLine & vbNewLine & "Click 'Cancel' to exit.", vbOKCancel, "Alert")
    If iReply = vbOK Then
        mainDoc.Save
    Else
        srErrorCheck = True
        Exit Function
    End If
End If

End Function

Private Function CreateErrorList(badStyles As String, arrStyleCount() As Variant) As String
Dim errorList As String
errorList = ""

'=====================Generate errors based on number of required elements found==================

'If Book Title = 0
If arrStyleCount(1) = 0 Then errorList = errorList & "**ERROR: No styled title detected." & _
    vbNewLine & vbNewLine

'If Book Title > 1
If arrStyleCount(1) > 1 Then errorList = errorList & "**ERROR: Too many title paragraphs detected." _
    & " Only 1 allowed." & vbNewLine & vbNewLine

'If Author Name = 0
If arrStyleCount(2) = 0 Then errorList = errorList & "**ERROR: No styled author name detected." _
    & vbNewLine & vbNewLine

'If ISBN = 0
If arrStyleCount(3) = 0 Then errorList = errorList & "**ERROR: No styled ISBN detected." _
    & vbNewLine & vbNewLine

'If CN > 1 and CT = 0 (already fixed in FixSoloCNs sub)
If arrStyleCount(4) > 0 And arrStyleCount(5) = 0 Then errorList = errorList & _
    "**WARNING: Chap Number (cn) cannot be the main heading for" & vbNewLine _
    & vbTab & "a chapter. Every chapter must start with Chapter Title (ct)" & vbNewLine _
    & vbTab & "style. Chap Number (cn) paragraphs have been converted to the" & vbNewLine _
    & vbTab & "Chap Title (ct) style." & vbNewLine & vbNewLine

'If no chapter opening paragraphs (CN, CT, or CTNP)
If arrStyleCount(4) = 0 And arrStyleCount(5) = 0 And arrStyleCount(6) = 0 Then errorList = errorList _
    & "**ERROR: No tagged chapter openers detected. If your book does" & vbNewLine _
    & vbTab & "not have chapter openers, use the Chap Title Nonprinting" & vbNewLine _
    & vbTab & "(ctnp) style at the start of each section." & vbNewLine & vbNewLine

'If CN > CT and CT > 0 (i.e., Not a CT for every CN)
If arrStyleCount(4) > arrStyleCount(5) And arrStyleCount(5) > 0 Then errorList = errorList & _
    "**ERROR: More Chap Number (cn) paragraphs than Chap Title (ct)" & vbNewLine _
    & vbTab & "paragraphs found. Each Chap Number (cn) paragraph MUST be" & vbNewLine _
    & vbTab & "followed by a Chap Title (ct) paragraph." & vbNewLine & vbNewLine

'If Imprint line = 0
If arrStyleCount(7) = 0 Then errorList = errorList & "**ERROR: No styled Imprint Line detected." _
    & vbNewLine & vbNewLine

'If Imprint Lline > 1
If arrStyleCount(7) > 1 Then errorList = errorList & "**ERROR: Too many Imprint Line paragraphs" _
    & " detected. Only 1 allowed." & vbNewLine & vbNewLine

'If only CTs (either originally or converted by macro) check for a page break before
If (arrStyleCount(4) > 0 And arrStyleCount(5) = 0) Or (arrStyleCount(4) = 0 And arrStyleCount(5) > 0) _
    Then errorList = errorList & CheckPrevStyle("Chap Title (ct)", "Page Break (pb)")

'If only CTNP, check for a page break before
If arrStyleCount(4) = 0 And arrStyleCount(5) = 0 And arrStyleCount(6) > 0 Then errorList = errorList _
    & CheckPrevStyle("Chap Title Nonprinting (ctnp)", "Page Break (pb)")

'If CNs are >= CTs, check for page break before CNs
If arrStyleCount(4) >= arrStyleCount(5) Then errorList = errorList & CheckPrevStyle("Chap Number (cn)", _
    "Page Break (pb)")

'If CNs and CTs, check that CN always comes before CT
If arrStyleCount(4) >= arrStyleCount(5) And arrStyleCount(5) <> 0 Then errorList = errorList _
    & CheckPrevStyle("Chap Title (ct)", "Chap Number (cn)")

'Check that only heading styles follow page breaks
errorList = errorList & CheckAfterPB

'Add bad styles to error message
    errorList = errorList & badStyles

If errorList <> "" Then
    errorList = errorList & vbNewLine & "If you have any questions about how to handle these errors, " & vbNewLine & _
        "please contact workflows@macmillan.com." & vbNewLine
End If

'Debug.Print errorList

CreateErrorList = errorList

End Function
Private Function GetText(styleName As String)
Dim fString As String
Dim fCount As Integer

fCount = 0

'Move selection to start of document
Selection.HomeKey Unit:=wdStory

    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Style = ActiveDocument.Styles(styleName)
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

Do While Selection.Find.Execute = True And fCount < 1000            'fCount < 1000 so we don't get an infinite loop
    fCount = fCount + 1
    
    'If paragraph return exists in selection, don't select last character (the last paragraph retunr)
    If InStr(Selection.Text, Chr(13)) > 0 Then
        Selection.MoveEnd Unit:=wdCharacter, Count:=-1
    End If
    
    'Assign selected text to variable
    fString = fString & Selection.Text & vbNewLine
    
    'If the next character is a paragraph return, add that to the selection
    'Otherwise the next Find will just select the same text with the paragraph return
    Selection.MoveEndWhile Cset:=Chr(13), Count:=wdForward
Loop

'Move selection back to start of document
Selection.HomeKey Unit:=wdStory

If fCount = 0 Then
    GetText = ""
Else
    GetText = fString
End If

End Function
Function CheckPrevStyle(findStyle As String, prevStyle As String)
Dim jString As String
Dim jCount As Integer
Dim pageNum As Integer

jCount = 0
jString = ""

'Move selection to start of document
Selection.HomeKey Unit:=wdStory

'select paragraph with that style
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Style = ActiveDocument.Styles(findStyle)
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

Do While Selection.Find.Execute = True And jCount < 1000            'jCount < 1000 so we don't get an infinite loop
    jCount = jCount + 1
    
    'select preceding paragraph
    Selection.Previous(Unit:=wdParagraph, Count:=1).Select
    pageNum = Selection.Information(wdActiveEndPageNumber)
    
        'Check if preceding paragraph style is correct
        If Selection.Style <> prevStyle Then
            jString = jString & "**ERROR: Missing or incorrect " & prevStyle & " style on page " & pageNum & "." & vbNewLine & vbNewLine
        End If
        
        'If you're searching for a page break before, also check if manual page break is in paragraph
        If prevStyle = "Page Break (pb)" Then
            If InStr(Selection.Text, Chr(12)) = 0 Then
                jString = jString & "**ERROR: Missing manual page break on page " & pageNum & "." & vbNewLine & vbNewLine
            End If
        End If
        
    'Debug.Print jString
    
    'move the selection back to original paragraph, so it won't be
    'selected again on next search
    Selection.Next(Unit:=wdParagraph, Count:=1).Select
Loop

'Debug.Print jString

CheckPrevStyle = jString

'Move selection back to start of document
Selection.HomeKey Unit:=wdStory

End Function
Function CheckAfterPB()
Dim kString As String
Dim kCount As Integer
Dim pageNumK As Integer
Dim nextStyle As String

kCount = 0
kString = ""

'Move selection to start of document
Selection.HomeKey Unit:=wdStory

'select paragraph with that style
    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Style = ActiveDocument.Styles("Page Break (pb)")
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

Do While Selection.Find.Execute = True And kCount < 1000            'jCount < 1000 so we don't get an infinite loop
    kCount = kCount + 1
    
    'select preceding paragraph
    Selection.Next(Unit:=wdParagraph, Count:=1).Select
    pageNumK = Selection.Information(wdActiveEndPageNumber)
    
        'Check if preceding paragraph style is correct
        If Selection.Style <> "Chap Title (ct)" And _
            Selection.Style <> "Chap Number (cn)" And _
            Selection.Style <> "Chap Title Nonprinting (ctnp)" And _
            Selection.Style <> "Halftitle Book Title (htit)" And _
            Selection.Style <> "Titlepage Book Title (tit)" And _
            Selection.Style <> "Copyright Text single space (crtx)" And _
            Selection.Style <> "Copyright Text double space (crtxd)" And _
            Selection.Style <> "Dedication (ded)" And _
            Selection.Style <> "Ad Card Main Head (acmh)" And _
            Selection.Style <> "Ad Card List of Titles (acl)" And _
            Selection.Style <> "Part Title (pt)" And _
            Selection.Style <> "Part Number (pn)" And _
            Selection.Style <> "Front Sales Title (fst)" And _
            Selection.Style <> "Front Sales Quote (fsq)" And _
            Selection.Style <> "Front Sales Quote NoIndent (fsq1)" And _
            Selection.Style <> "Epigraph - non-verse (epi)" And _
            Selection.Style <> "Epigraph - verse (epiv)" And _
            Selection.Style <> "FM Head (fmh)" And _
            Selection.Style <> "Illustration holder (ill)" And _
            Selection.Style <> "Page Break (pb)" Then
                nextStyle = Selection.Style
                kString = kString & "**ERROR: Missing or incorrect Page Break or " & nextStyle & " style on page " & pageNumK & "." & vbNewLine & vbNewLine
        End If
                
    'Debug.Print kString
    
    'move the selection back to original paragraph, so it won't be
    'selected again on next search
    Selection.Previous(Unit:=wdParagraph, Count:=1).Select
Loop

'Debug.Print kString

CheckAfterPB = kString

'Move selection back to start of document
Selection.HomeKey Unit:=wdStory

End Function
Private Sub DeleteContentControlPC()
Dim cc As ContentControl

For Each cc In ActiveDocument.ContentControls
    cc.Delete
Next
End Sub
Private Function FixTrackChanges() As Boolean
Dim N As Long
Dim oComments As Comments
Set oComments = ActiveDocument.Comments

FixTrackChanges = True

Application.DisplayAlerts = False

'Turn off track changes
ActiveDocument.TrackRevisions = False

'See if there are tracked changes or comments in document
On Error Resume Next
Selection.HomeKey Unit:=wdStory   'start search at beginning of doc
WordBasic.NextChangeOrComment       'search for a tracked change or comment. error if none are found.

'If there are changes, ask user if they want macro to accept changes or cancel
If Err = 0 Then
    If MsgBox("Bookmaker doesn't like comments or tracked changes, but it appears that you have some in your document." _
        & vbCr & vbCr & "Click OK to ACCEPT ALL CHANGES and DELETE ALL COMMENTS right now and continue with the Bookmaker Requirements Check." _
        & vbCr & vbCr & "Click CANCEL to stop the Bookmaker Requirements Check and deal with the tracked changes and comments on your own.", _
        273, "Are those tracked changes I see?") = vbCancel Then           '273 = vbOkCancel(1) + vbCritical(16) + vbDefaultButton2(256)
            FixTrackChanges = False
            Exit Function
    Else 'User clicked OK, so accept all tracked changes and delete all comments
        ActiveDocument.AcceptAllRevisions
        For N = oComments.Count To 1 Step -1
            oComments(N).Delete
        Next N
        Set oComments = Nothing
    End If
End If

On Error GoTo 0
Application.DisplayAlerts = True

End Function
Private Function BadTorStyles() As String
'Called from GoodBadStyles sub if torDOTcom parameter is set to True.

Dim paraStyle As String
Dim activeParaCount As Integer

Dim arrTorStyles() As String
ReDim arrTorStyles(1 To 76)

Dim intBadCount As Integer
Dim activeParaRange As Range
Dim pageNumber As Integer

Dim N As Integer
Dim M As Integer
Dim strBadStyles As String

'List of styles approved for use in Tor.com automated workflow
'Organized by approximate frequency in manuscripts (most freq at top)

arrTorStyles(1) = "Text - Standard (tx)"
arrTorStyles(2) = "Text - Std No-Indent (tx1)"
arrTorStyles(3) = "Chap Title (ct)"
arrTorStyles(4) = "Chap Number (cn)"
arrTorStyles(5) = "Chap Opening Text No-Indent (cotx1)"
arrTorStyles(6) = "Page Break (pb)"
arrTorStyles(7) = "Space Break (#)"
arrTorStyles(8) = "Space Break with Ornament (orn)"
arrTorStyles(9) = "Titlepage Author Name (au)"
arrTorStyles(10) = "Titlepage Book Subtitle (stit)"
arrTorStyles(11) = "Titlepage Book Title (tit)"
arrTorStyles(12) = "Titlepage Cities (cit)"
arrTorStyles(13) = "Titlepage Imprint Line (imp)"
arrTorStyles(14) = "Copyright Text double space (crtxd)"
arrTorStyles(15) = "Copyright Text single space (crtx)"
arrTorStyles(16) = "Extract (ext)"
arrTorStyles(17) = "Extract Head (exth)"
arrTorStyles(18) = "Extract-No Indent (ext1)"
arrTorStyles(19) = "Halftitle Book Title (htit)"
arrTorStyles(20) = "Illustration holder (ill)"
arrTorStyles(21) = "Illustration Source (is)"
arrTorStyles(22) = "Part Number (pn)"
arrTorStyles(23) = "Part Title (pt)"
arrTorStyles(24) = "About Author Text Head (atah)"
arrTorStyles(25) = "About Author Text (atatx)"
arrTorStyles(26) = "About Author Text No-Indent (atatx1)"
arrTorStyles(27) = "Ad Card List of Titles (acl)"
arrTorStyles(28) = "Ad Card Main Head (acmh)"
arrTorStyles(29) = "Ad Card Subhead (acsh)"
arrTorStyles(30) = "Text - Standard Space After (tx#)"
arrTorStyles(31) = "Text - Standard Space Around (#tx#)"
arrTorStyles(32) = "Text - Standard Space Before (#tx)"
arrTorStyles(33) = "Text - Std No-Indent Space After (tx1#)"
arrTorStyles(34) = "Text - Std No-Indent Space Around (#tx1#)"
arrTorStyles(35) = "Text - Std No-Indent Space Before (#tx1)"
arrTorStyles(36) = "Chap Opening Text No-Indent Space After (cotx1#)"
arrTorStyles(37) = "Dedication (ded)"
arrTorStyles(38) = "Dedication Author (dedau)"
arrTorStyles(39) = "Epigraph - non-verse (epi)"
arrTorStyles(40) = "Epigraph - verse (epiv)"
arrTorStyles(41) = "Epigraph Source (eps)"
arrTorStyles(42) = "Chap Epigraph Source (ceps)"
arrTorStyles(43) = "Chap Epigraph—non-verse (cepi)"
arrTorStyles(44) = "Chap Epigraph—verse (cepiv)"
arrTorStyles(45) = "Chap Title Nonprinting (cnp)"
arrTorStyles(46) = "FM Epigraph - non-verse (fmepi)"
arrTorStyles(47) = "FM Epigraph - verse (fmepiv)"
arrTorStyles(48) = "FM Epigraph Source (fmeps)"
arrTorStyles(49) = "FM Head (fmh)"
arrTorStyles(50) = "FM Subhead (fmsh)"
arrTorStyles(51) = "FM Text (fmtx)"
arrTorStyles(52) = "FM Text No-Indent (fmtx1)"
arrTorStyles(53) = "FM Text No-Indent Space After (fmtx1#)"
arrTorStyles(54) = "FM Text No-Indent Space Around (#fmtx1#)"
arrTorStyles(55) = "FM Text No-Indent Space Before (#fmtx1)"
arrTorStyles(56) = "FM Text Space After (fmtx#)"
arrTorStyles(57) = "FM Text Space Around (#fmtx#)"
arrTorStyles(58) = "FM Text Space Before (#fmtx)"
arrTorStyles(59) = "Front Sales Quote (fsq)"
arrTorStyles(60) = "Front Sales Quote NoIndent (fsq1)"
arrTorStyles(61) = "Front Sales Quote Source (fsqs)"
arrTorStyles(62) = "Front Sales Title (fst)"
arrTorStyles(63) = "Space Break with ALT Ornament (orn2)"
arrTorStyles(64) = "Space Break - 1-Line (ls1)"
arrTorStyles(65) = "Space Break - 2-Line (ls2)"
arrTorStyles(66) = "Space Break - 3-Line (ls3)"
arrTorStyles(67) = "Text - Computer Type (com)"
arrTorStyles(68) = "Text - Computer Type No-Indent (com1)"
arrTorStyles(69) = "Text - Standard ALT (atx)"
arrTorStyles(70) = "Text - Std No-Indent ALT (atx1)"
arrTorStyles(71) = "Caption (cap)"
arrTorStyles(72) = "Titlepage Contributor Name (con)"
arrTorStyles(73) = "Titlepage Translator Name (tran)"
arrTorStyles(74) = "Chap Ornament (corn)"
arrTorStyles(75) = "Chap Ornament ALT (corn2)"
arrTorStyles(76) = "Chap Opening Text (cotx)"

activeParaCount = ActiveDocument.paragraphs.Count

For N = 1 To activeParaCount
    intBadCount = 0
    paraStyle = ActiveDocument.paragraphs(N).Style
    
    If Right(paraStyle, 1) = ")" Then
    
        For M = LBound(arrTorStyles()) To UBound(arrTorStyles())
            If paraStyle <> arrTorStyles(M) Then
                intBadCount = intBadCount + 1
            Else
                Exit For
            End If
        Next M
    
        If intBadCount = UBound(arrTorStyles()) Then
            Set activeParaRange = ActiveDocument.paragraphs(N).Range
            pageNumber = activeParaRange.Information(wdActiveEndPageNumber)
            strBadStyles = strBadStyles & "** ERROR: Non-Tor.com style on page " & pageNumber _
                & " (Paragraph " & N & "): " & vbTab & paraStyle & vbNewLine & vbNewLine
        End If
    
    End If


Next N

Debug.Print strBadStyles

BadTorStyles = strBadStyles

End Function
Private Function CountReqdStyles() As Variant
Dim arrStyleName(1 To 7) As String                      ' Declare number of items in array
Dim intStyleCount() As Variant
ReDim intStyleCount(1 To 7) As Variant
Dim A As Long
Dim xCount As Integer

arrStyleName(1) = "Titlepage Book Title (tit)"
arrStyleName(2) = "Titlepage Author Name (au)"
arrStyleName(3) = "span ISBN (isbn)"
arrStyleName(4) = "Chap Number (cn)"
arrStyleName(5) = "Chap Title (ct)"
arrStyleName(6) = "Chap Title Nonprinting (ctnp)"
arrStyleName(7) = "Titlepage Imprint Line (imp)"

For A = 1 To UBound(arrStyleName())
    xCount = 0
    
    With ActiveDocument.Range.Find
        .ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Style = ActiveDocument.Styles(arrStyleName(A))
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    Do While .Execute(Forward:=True) = True And xCount < 100   'xCount < 100 to precent infinite loop, especially if content controls in title or author blocks
        intStyleCount(A) = intStyleCount(A) + 1
        xCount = xCount + 1
    Loop
    End With
Next

            
'------------Exit Sub if exactly 10 Titles styled, suggests hidden content controls-----
If intStyleCount(1) = 100 Then
    
    MsgBox "Something went wrong!" & vbCr & vbCr & "It looks like you might have content controls (form fields or drop downs) in your document, but Word for Mac doesn't play nicely with these." _
    & vbCr & vbCr & "Try running this macro on a PC or contact workflows@macmillan.com for assistance.", vbCritical, "OH NO!!"
    Exit Function
    
End If

'For A = 1 To UBound(arrStyleName())
 '   Debug.Print arrStyleName(A) & ": " & intStyleCount(A) & vbNewLine
'Next A

CountReqdStyles = intStyleCount()

End Function
Private Sub FixSoloCNs()

'Move selection to start of document
Selection.HomeKey Unit:=wdStory

'Find paras styles as CN and change to CT style
    Selection.Find.ClearFormatting
    Selection.Find.Style = ActiveDocument.Styles("Chap Number (cn)")
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Style = ActiveDocument.Styles("Chap Title (ct)")
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
End Sub

Private Function GetMetadata() As String
Dim styleNameB(4) As String         ' must declare number of items in array here
Dim bString(4) As String            ' and here
Dim b As Integer
Dim strTitleData As String

styleNameB(1) = "Titlepage Book Title (tit)"
styleNameB(2) = "Titlepage Author Name (au)"
styleNameB(3) = "span ISBN (isbn)"
styleNameB(4) = "Titlepage Imprint Line (imp)"

For b = 1 To UBound(styleNameB())
    bString(b) = GetText(styleNameB(b))
Next b

strTitleData = "** Title **" & vbNewLine & _
            bString(1) & vbNewLine & _
            "** Author **" & vbNewLine & _
            bString(2) & vbNewLine & _
            "** ISBN **" & vbNewLine & _
            bString(3) & vbNewLine & _
            "** Imprint **" & vbNewLine & _
            bString(4) & vbNewLine
            
'Debug.Print strTitleData

GetMetadata = strTitleData

End Function

Private Function IllustrationsList() As String
Dim cString(1000) As String             'Max number of illustrations. Could be lower than 1000.
Dim cCount As Integer
Dim pageNumberC As Integer
Dim strFullList As String
Dim N As Integer

cCount = 0

'Move selection to start of document
Selection.HomeKey Unit:=wdStory

    Selection.Find.ClearFormatting
    With Selection.Find
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Style = ActiveDocument.Styles("Illustration holder (ill)")
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

Do While Selection.Find.Execute = True And cCount < 1000            'cCount < 1000 so we don't get an infinite loop
    cCount = cCount + 1
    pageNumberC = Selection.Information(wdActiveEndPageNumber)
    
    'If paragraph return exists in selection, don't select last character (the last paragraph return)
    If InStr(Selection.Text, Chr(13)) > 0 Then
        Selection.MoveEnd Unit:=wdCharacter, Count:=-1
    End If
    
    cString(cCount) = "Page " & pageNumberC & ": " & Selection.Text
    
    'If the next character is a paragraph return, add that to the selection
    'Otherwise the next Find will just select the same text with the paragraph return
    Selection.MoveEndWhile Cset:=Chr(13), Count:=wdForward
    
Loop

'Move selection back to start of document
Selection.HomeKey Unit:=wdStory

If cCount > 1000 Then
    MsgBox "You have more than 1,000 illustrations tagged in your manuscript." & vbNewLine & _
    "Please contact workflows@macmillan.com to complete your illustration list."
End If

If cCount = 0 Then
    cCount = 1
    cString(1) = "no illustrations detected" & vbNewLine
End If

For N = 1 To cCount
    strFullList = strFullList & cString(N)
Next N

'Debug.Print strFullList

IllustrationsList = strFullList


End Function

Private Sub CreateReport(errorList As String, metadata As String, illustrations As String, goodStyles As String, suffix As String)
'Create report file
Dim activeRng As Range
Dim activeDoc As Document
Set activeDoc = ActiveDocument
Set activeRng = ActiveDocument.Range
Dim activeDocName As String
Dim activeDocPath As String
Dim reqReportDoc As String
Dim reqReportDocAlt As String
Dim fnum As Integer
Dim TheOS As String
TheOS = System.OperatingSystem

'activeDocName below works for .doc and .docx
activeDocName = Left(activeDoc.Name, InStrRev(activeDoc.Name, ".doc") - 1)
activeDocPath = Replace(activeDoc.Path, activeDoc.Name, "")

'create text file
reqReportDoc = activeDocPath & activeDocName & "_" & suffix & ".txt"

''''for 32 char Mc OS bug- could check if this is Mac OS too < PART 1
If Not TheOS Like "*Mac*" Then                      'If Len(activeDocName) > 18 Then        (legacy, does not take path into account)
    reqReportDoc = activeDocPath & "\" & activeDocName & "_" & suffix & ".txt"
Else
    Dim placeholdDocName As String
    placeholdDocName = "filenamePlacehold_Report.txt"
    reqReportDocAlt = reqReportDoc
    reqReportDoc = "Macintosh HD:private:tmp:" & placeholdDocName
End If
'''end ''''for 32 char Mc OS bug part 1

'set and open file for output
Dim e As Integer

fnum = FreeFile()
Open reqReportDoc For Output As fnum
If errorList = "" Then
    Print #fnum, vbCr
    Print #fnum, "                 CONGRATULATIONS! YOU PASSED!" & vbCr
    Print #fnum, " But you're not done yet. Please check the info listed below." & vbCr
    Print #fnum, vbCr

Else
    Print #fnum, vbCr
    Print #fnum, "                             OOPS!" & vbCr
    Print #fnum, "     Problems were found with the styles in your document." & vbCr
    Print #fnum, vbCr
    Print #fnum, vbCr
    Print #fnum, "--------------------------- ERRORS ---------------------------" & vbCr
    Print #fnum, errorList
    Print #fnum, vbCr
    Print #fnum, vbCr
End If
    Print #fnum, "--------------------------- METADATA -------------------------" & vbCr
    Print #fnum, "If any of the information below is wrong, please fix the" & vbCr
    Print #fnum, "associated styles in the manuscript." & vbCr
    Print #fnum, vbCr
    Print #fnum, metadata
    Print #fnum, vbCr
    Print #fnum, vbCr
    Print #fnum, "----------------------- ILLUSTRATION LIST ---------------------" & vbCr
    
    If illustrations <> "no illustrations detected" & vbNewLine Then
        Print #fnum, "Verify that this list of illustrations includes only the file" & vbCr
        Print #fnum, "names of your illustrations. Be sure to place these files in" & vbCr
        Print #fnum, "the submitted_images folder BEFORE you run the bookmaker tool." & vbCr
        Print #fnum, vbCr
    End If
    
    Print #fnum, illustrations
    Print #fnum, vbCr
    Print #fnum, vbCr
    Print #fnum, "----------------------- MACMILLAN STYLES IN USE --------------------" & vbCr
    Print #fnum, goodStyles

Close #fnum

''''for 32 char Mc OS bug-<PART 2
If reqReportDocAlt <> "" Then
Name reqReportDoc As reqReportDocAlt
End If
''''END for 32 char Mac OS bug-<PART 2

'----------------open Bookmaker Report for user once it is complete--------------------------.
Dim Shex As Object

If Not TheOS Like "*Mac*" Then
   Set Shex = CreateObject("Shell.Application")
   Shex.Open (reqReportDoc)
Else
    MacScript ("tell application ""TextEdit"" " & vbCr & _
    "open " & """" & reqReportDocAlt & """" & " as alias" & vbCr & _
    "activate" & vbCr & _
    "end tell" & vbCr)
End If
End Sub

