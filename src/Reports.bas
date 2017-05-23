Attribute VB_Name = "Reports"
' ====== PURPOSE =================
' Checks that manuscript styles follow Macmillan best practices
' If no Macmillan styles are used, just produces a list of styles in use

' ====== DEPENDENCIES ============
' 1. Manuscript must be styled with Macmillan custom styles to generate full report.
' 2. Requires ProgressBar userform
' 3. Requires MacroHelpers be installed in same template.


Option Explicit
Option Base 1

Private m_strVersion As String

Sub BookmakerReqs()
    Call MakeReport(torDOTcom:=True)
End Sub


Sub MacmillanStyleReport()
    Call MakeReport(torDOTcom:=False)
End Sub


Public Function GetStyleVersion() As String

  Dim strStyleVersion As String
  If Utils.DocPropExists(objDoc:=activeDoc, PropName:="Version") = True Then
    strStyleVersion = activeDoc.CustomDocumentProperties("Version").Value
  Else
    strStyleVersion = vbNullString
  End If

End Function

Private Sub MakeReport(torDOTcom As Boolean)
    '-----------------------------------------------------------
    
    
    '=================================================
    '''''              Timer Start                  '|
    'Dim StartTime As Double                         '|
    'Dim SecondsElapsed As Double                    '|
                                                    '|
    '''''Remember time when macro starts            '|
    'StartTime = Timer                               '|
    '=================================================
    

  
  ' ======= Run startup checks ========
  ' True means a check failed (e.g., doc protection on)
  If MacroHelpers.StartupSettings() = True Then
    Call Cleanup
    Exit Sub
  End If
  
'  Dim blnOldStartStyles As Boolean
'  m_strVersion = GetStyleVersion()
'  If m_strVersion = vbNullString Then
'    blnOldStartStyles = True
'  Else
'    blnOldStartStyles = False
'  End If
    
  '--------Progress Bar------------------------------
  'Percent complete and status for progress bar (PC) and status bar (Mac)
  'Requires ProgressBar custom UserForm and Class
  Dim sglPercentComplete As Single
  Dim strStatus As String
  Dim strTitle As String
  
  'First status shown will be randomly pulled from array, for funzies
  Dim funArray() As String
  ReDim funArray(1 To 10)      'Declare bounds of array here
  
  If torDOTcom = True Then
    funArray(1) = "* Is this thing on?..."
    funArray(2) = "* Are we there yet?..."
    funArray(3) = "* Zapping space invaders..."
    funArray(4) = "* Leaping over tall buildings in a single bound..."
    funArray(5) = "* Taking a quick nap..."
    funArray(6) = "* Taking the stairs..."
    funArray(7) = "* Partying like it's 1999..."
    funArray(8) = "* Waiting in line at Shake Shack..."
    funArray(9) = "* Revving engines..."
    funArray(10) = "* Thanks for running the Bookmaker Macro!"
  Else
    funArray(1) = "* Now is the winter of our discontent, made glorious summer by these Word Styles..."
    funArray(2) = "* What's in a name? Word Styles by any name would smell as sweet..."
    funArray(3) = "* A horse! A horse! My Word Styles for a horse!"
    funArray(4) = "* Be not afraid of Word Styles. Some are born with Styles, some achieve Styles, and some have Styles thrust upon 'em..."
    funArray(5) = "* All the world's a stage, and all the Word Styles merely players..."
    funArray(6) = "* To thine own Word Styles be true, and it must follow, as the night the day, thou canst not then be false to any man..."
    funArray(7) = "* To Style, or not to Style: that is the question..."
    funArray(8) = "* Word Styles, Word Styles! Wherefore art thou Word Styles?..."
    funArray(9) = "* Some Cupid kills with arrows, some with Word Styles..."
    funArray(10) = "* What light through yonder window breaks? It is the east, and Word Styles are the sun..."
  End If
  
  Dim X As Integer
    
  'Rnd returns random number between (0,1], rest of expression is to return an integer (1,10)
  Randomize           'Sets seed for Rnd below to value of system timer
  X = Int(UBound(funArray()) * Rnd()) + 1
  
  'DebugPrint x
  ' If only creating an epub, just run style report since only difference now
  ' is Bkmkr checks for non-Bookmaker styles.
  If torDOTcom = True Then
    strTitle = "Bookmaker Requirements Macro"
    Dim strEpubMsg As String
    strEpubMsg = "Hi there! Are you creating a PRINT PDF with Bookmaker?" & _
      vbNewLine & vbNewLine & _
      "If you are creating a PRINT PDF, click YES." & vbNewLine & _
      "If you are ONLY creating an EPUB, click NO."
    If MsgBox(strEpubMsg, vbYesNo) = vbNo Then
      torDOTcom = False
    End If
  Else
    strTitle = "Macmillan Style Report"
  End If
    
  sglPercentComplete = 0.02
  strStatus = funArray(X)
  
  Dim oProgressBkmkr As ProgressBar
  Set oProgressBkmkr = New ProgressBar    ' Triggers Initialize

  oProgressBkmkr.Title = strTitle
  Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)

' ------------check for endnotes and footnotes---------------------------------
  Dim colStories As Collection
  Set colStories = MacroHelpers.ActiveStories
  Dim varStory As Variant
  Dim currentStory As WdStoryType


' -------- validate section-start styles --------------------------------------
  Dim strSectionStartWarnings As String
  strSectionStartWarnings = SectionStartRules()
  
' -------- Remove formatting from CN paragraphs -----------------------------------------------
  Call ChapNumCleanUp

  '-------remove "span ISBN (isbn)" style from letters, spaces, parens, etc.-------------------
  '-------because it should just be applied to the isbn numerals and hyphens-------------------
  Call ISBNcleanup
  
' -------------- Clean up page break characters -------------------------------
  Call MacroHelpers.PageBreakCleanup
    
  '--------Get title/author/isbn/imprint text from document-----------
  sglPercentComplete = 0.11
  Application.ScreenUpdating = True
  strStatus = "* Getting book metadata from manuscript..." & vbCr & strStatus
  
  Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)
  
  Dim strMetadata As String
  strMetadata = GetMetadata
    
  '-------------------Get Illustrations List from Document-----------
  sglPercentComplete = 0.15
  strStatus = "* Getting list of illustrations..." & vbCr & strStatus
  
  Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)
  
  Dim strIllustrationsList As String
  strIllustrationsList = IllustrationsList
      
  '-------------------Get list of good and bad styles from document---------
  sglPercentComplete = 0.18
  strStatus = "* Getting list of styles in use..." & vbCr & strStatus
  
  Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)
  
  Dim arrGoodBadStyles() As Variant
  Dim strGoodStylesList As String
  Dim strBadStylesList As String
              
  'returns array with 2 elements, 1: good styles list, 2: bad styles list
  arrGoodBadStyles = GoodBadStyles(Tor:=torDOTcom, ProgressBar:=oProgressBkmkr, Status:=strStatus, ProgTitle:=strTitle, _
      Stories:=arrStories)
  strGoodStylesList = arrGoodBadStyles(1)
  'DebugPrint strGoodStylesList
  strBadStylesList = arrGoodBadStyles(2)
        
    'Error checking: if no good styles are in use, just return list of all styles in use, not other checks
    Dim blnTemplateUsed As Boolean
    Dim strSearchPattern As String
    ' Searching for "Footnote Text" or "Endnote Text" followed by page number, then
    ' followed by anything NOT including a close bracket. If there are other Mac styles
    ' it won't select the whole string
    
    strSearchPattern = "[EF]{1}[dnot]{4}[eot]{2,} Text -- p. [0-9]{1,}[!\)]{1,}"
    
    If strGoodStylesList = vbNullString Then
        blnTemplateUsed = False
    ' Test if good styles are just Endnote Text and Footnote Text
    ElseIf PatternMatch(SearchPattern:=strSearchPattern, SearchText:=strGoodStylesList, WholeString:=True) = True Then
        blnTemplateUsed = False
    Else
        blnTemplateUsed = True
    End If
    
    'If template not used, just returns list of styles in use
    If blnTemplateUsed = False Then
        strGoodStylesList = StylesInUse(ProgressBar:=oProgressBkmkr, Status:=strStatus, ProgTitle:=strTitle, Stories:=arrStories)
        strBadStylesList = ""
    End If
    
    '-------------------Create list of errors----------------------------
    sglPercentComplete = 0.98
    strStatus = "* Checking styles for errors..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)
    
    Dim strErrorList As String
    
    If blnTemplateUsed = True Then
        strErrorList = CreateErrorList(badStyles:=strBadStylesList, strSecStWarnings:=strSectionStartWarnings, blnTor:=torDOTcom)
        'strErrorList = "testing"
    Else
        strErrorList = ""
    End If
    
    '------Create Report Text-------------------------------
    sglPercentComplete = 0.99
    strStatus = "* Creating report file..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)
    
    Dim strReportText As String
    strReportText = CreateReportText(blnTemplateUsed, strErrorList, strMetadata, strIllustrationsList, strGoodStylesList)
    
    
    ' Create Report File -------------------------------------
    Dim strSuffix As String
    If torDOTcom = True Then
        strSuffix = "BookmakerReport" ' suffix for the report file
    Else
        strSuffix = "StyleReport"
    End If
    
    Call MacroHelpers.CreateTextFile(strText:=strReportText, suffix:=strSuffix)
    
    '-------------Go back to original settings-----------------
    sglPercentComplete = 1
    strStatus = "* Finishing up..." & vbCr & strStatus
    
    Call UpdateBarAndWait(Bar:=oProgressBkmkr, Status:=strStatus, Percent:=sglPercentComplete)
    
    Call Cleanup
    
    Unload oProgressBkmkr
    
    '============================================================================
    '----------------------Timer End-------------------------------------------
    ''''Determine how many seconds code took to run
      'SecondsElapsed = Round(Timer - StartTime, 2)
    
    ''''Notify user in seconds
      'DebugPrint "This code ran successfully in " & SecondsElapsed & " seconds"
    '============================================================================

End Sub

Private Function GoodBadStyles(Tor As Boolean, ProgressBar As ProgressBar, Status As String, ProgTitle As String) As Variant
    'Creates a list of Macmillan styles in use
    'And a separate list of non-Macmillan styles in use
    
    Dim TheOS As String
    TheOS = System.OperatingSystem
    Dim sglPercentComplete As Single
    Dim strStatus As String
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
    
    Dim colStories As Collection
    Set colStories = MacroHelpers.ActiveStories
    
    Dim varStory As Variant
    Dim currentStory As WdStoryType
    
    'Alter built-in Normal (Web) style temporarily (later, maybe forever?)
    activeDoc.Styles("Normal (Web)").NameLocal = "_"
    
    '----------Collect all styles being used-------------------------------
    styleGoodCount = 0
    styleBadCount = 0
    styleBadOverflow = False
    activeParaCount = activeDoc.Paragraphs.Count
    For J = 1 To activeParaCount
        
        'All Progress Bar statements for PC only because won't run modeless on Mac
        If J Mod 100 = 0 Then
        
            'Percent complete and status for progress bar (PC) and status bar (Mac)
            sglPercentComplete = (((J / activeParaCount) * 0.45) + 0.18)
            strStatus = "* Checking paragraph " & J & " of " & activeParaCount & " for Macmillan styles..." & _
                        vbCr & Status
            
            'DebugPrint sglPercentComplete
            Call UpdateBarAndWait(Bar:=ProgressBar, Status:=strStatus, Percent:=sglPercentComplete)
        End If
        

        For Each varStory In colStories
          currentStory = varStory
            If J <= activeDoc.StoryRanges(currentStory).Paragraphs.Count Then
                paraStyle = activeDoc.StoryRanges(currentStory).Paragraphs(J).Style
                Set activeParaRange = activeDoc.StoryRanges(currentStory).Paragraphs(J).Range
                pageNumber = activeParaRange.Information(wdActiveEndPageNumber)                 'alt: (wdActiveEndAdjustedPageNumber)
                    
                'If InStrRev(paraStyle, ")", -1, vbTextCompare) Then        'ALT calculation to "Right", can speed test
                If Right(paraStyle, 1) = ")" Then
CheckGoodStyles:
                    For K = 1 To styleGoodCount
                        'DebugPrint Left(stylesGood(K), InStrRev(stylesGood(K), " --") - 1)
                        ' "Left" function because now stylesGood includes page number, so won't match paraStyle
                        If paraStyle = Left(stylesGood(K), InStrRev(stylesGood(K), " --") - 1) Then
                        K = styleGoodCount                              'stylereport bug fix #1    v. 3.1
                            Exit For                                        'stylereport bug fix #1   v. 3.1
                        End If                                              'stylereport bug fix #1   v. 3.1
                    Next K
                    
                    If K = styleGoodCount + 1 Then
                        styleGoodCount = K
                        ReDim Preserve stylesGood(1 To styleGoodCount)
                        stylesGood(styleGoodCount) = paraStyle & " -- p. " & pageNumber
                    End If
                
                Else
                    
                    If paraStyle = "Endnote Text" Or paraStyle = "Footnote Text" Then
                        GoTo CheckGoodStyles
                    Else
                        For L = 1 To styleBadCount
                            'If paraStyle = stylesBad(L) Then Exit For                  'Not needed, since we want EVERY instance of bad style
                        Next L
                        If L > 100 Then                                                 ' Exits if more than 100 bad paragraphs
                                styleBadOverflow = True
                                stylesBad(100) = "** WARNING: More than 100 paragraphs with bad styles found." & vbNewLine & vbNewLine
                            Exit For
                        End If
                        If L = styleBadCount + 1 Then
                            styleBadCount = L
            
                            stylesBad(styleBadCount) = "** ERROR: Non-Macmillan style on page " & pageNumber & _
                                " (Paragraph " & J & "):  " & paraStyle & vbNewLine & vbNewLine
                        End If
                     End If
                End If
            End If
        Next A
    Next J
    
    Status = "* Checking paragraphs for Macmillan styles..." & vbCr & Status
    
    'Change Normal (Web) back (if you want to)
    activeDoc.Styles("Normal (Web),_").NameLocal = "Normal (Web)"
    
    ' DON'T sort styles alphabetically, per request from PE
'    'Sort good styles
'    If K <> 0 Then
'    ReDim Preserve stylesGood(1 To styleGoodCount)
'    WordBasic.SortArray stylesGood()
'    End If
    
    'Create single string for good styles
    Dim strGoodStyles As String
    
    If styleGoodCount = 0 Then
        strGoodStyles = ""
    Else
        For K = LBound(stylesGood()) To UBound(stylesGood())
            strGoodStyles = strGoodStyles & stylesGood(K) & vbCrLf
        Next K
    End If
    
    'DebugPrint strGoodStyles
    
    If styleBadCount > 0 Then
        'Create single string for bad styles
        Dim strBadStyles As String
        ReDim Preserve stylesBad(1 To styleBadCount)
        For L = LBound(stylesBad()) To UBound(stylesBad())
            strBadStyles = strBadStyles & stylesBad(L)
        Next L
    Else
        strBadStyles = ""
    End If
    
    'DebugPrint strBadStyles
    
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
    styleNameM(17) = "span strikethrough characters (str)"
    styleNameM(18) = "bookmaker keep together (kt)"
    styleNameM(19) = "span ISBN (isbn)"
    styleNameM(20) = "span symbols ital (symi)"
    styleNameM(21) = "span symbols bold (symb)"
    
    
    
    For M = 1 To UBound(styleNameM())
        
        'Percent complete and status for progress bar (PC) and status bar (Mac)
        sglPercentComplete = (((M / UBound(styleNameM())) * 0.13) + 0.63)
        strStatus = "* Checking for " & styleNameM(M) & " styles..." & vbCr & Status
    
        Call UpdateBarAndWait(Bar:=ProgressBar, Status:=strStatus, Percent:=sglPercentComplete)
        
        On Error GoTo ErrHandler
        
        'Move selection back to start of document
        Selection.HomeKey Unit:=wdStory
        
        'Need to do Selection.Find for char styles. Range.Find won't work.
        'We only need to find a style once to add it to the list
        'Search through the main text story here
        With Selection.Find
            .Style = activeDoc.Styles(styleNameM(M))
            .Wrap = wdFindContinue
            .Format = True
            .Execute
        End With
        
        If Selection.Find.Found = True Then
            charStyles = charStyles & styleNameM(M) & vbNewLine
        'Else not present in main text story
        Else
            ' So check if there are footnotes
            If activeDoc.Footnotes.Count > 0 Then
                'If there are footnotes, select the footnote text
                activeDoc.StoryRanges(wdFootnotesStory).Select
                'Search the new selection for the style
                With Selection.Find
                    .Style = activeDoc.Styles(styleNameM(M))
                    .Wrap = wdFindContinue
                    .Format = True
                    .Execute
                End With
            
                If Selection.Find.Found = True Then
                    charStyles = charStyles & styleNameM(M) & vbNewLine
                ' Else didn't find style in footnotes, check endnotes
                Else
                    GoTo CheckEndnotes
                End If
            Else
CheckEndnotes:
                ' Check if there are endnotes in the document
                If activeDoc.Endnotes.Count > 0 Then
                    ' If there are endnotes, select them
                    activeDoc.StoryRanges(wdEndnotesStory).Select
                    'Search the new selection for the style
                    With Selection.Find
                         .Style = activeDoc.Styles(styleNameM(M))
                         .Wrap = wdFindContinue
                         .Format = True
                         .Execute
                     End With
                        
                    If Selection.Find.Found = True Then
                        charStyles = charStyles & styleNameM(M) & vbNewLine
                    End If
                End If
            End If
        End If
NextLoop:
    Next M
    
    'DebugPrint charStyles
    
    Status = "* Checking character styles..." & vbCr & Status
    
    'Add character styles to Good styles list
    strGoodStyles = strGoodStyles & charStyles
    
    'If this is for the Tor.com Bookmaker toolchain, test if only those styles used
    Dim strTorBadStyles As String
    If Tor = True Then
        strTorBadStyles = BadTorStyles(ProgressBar2:=ProgressBar, StatusBar:=Status, ProgressTitle:=ProgTitle)
        strBadStyles = strBadStyles & strTorBadStyles
    End If
    
    'DebugPrint strGoodStyles
    'DebugPrint strBadStyles
    
    'If only good styles are Endnote Text and Footnote text, then the template is not being used
    
    
    'Add both good and bad styles lists to an array to pass back to original sub
    Dim arrFinalLists() As Variant
    ReDim arrFinalLists(1 To 2)
    
    arrFinalLists(1) = strGoodStyles
    arrFinalLists(2) = strBadStyles
    
    GoodBadStyles = arrFinalLists
    
    Exit Function
    
ErrHandler:
    'DebugPrint Err.Number & " : " & Err.Description
    If Err.Number = 5834 Or Err.Number = 5941 Then
        Resume NextLoop
    End If
    
End Function


Private Function CreateErrorList(badStyles As String, strSecStWarnings As String, arrStyleCount() As Variant, blnTor As Boolean) As String
    Dim errorList As String
    
    errorList = strSecStWarnings
    
    'Bookmaker file name validation
    If blnTor = True Then
        If CheckFileName = True Then errorList = errorList & _
            "**ERROR: Bookmaker can only accept file names that use" & vbNewLine & _
            "letters, numbers, hyphens, or underscores. Punctuation," & vbNewLine & _
            "spaces, and other special characters are not allowed." & vbNewLine & vbNewLine
    End If
  
    'Add bad styles to error message
    errorList = errorList & badStyles
    
    If errorList <> "" Then
        errorList = errorList & vbNewLine & "If you have any questions about how to handle these errors, " & vbNewLine & _
            "please contact workflows@macmillan.com." & vbNewLine
    End If
    
    'DebugPrint errorList
    
    CreateErrorList = errorList

End Function


Function CheckPrevStyle(findStyle As String, prevStyle As String) As String
    Dim jString As String
    Dim jCount As Integer
    Dim pageNum As Integer
    Dim intCurrentPara As Integer
    
    Application.ScreenUpdating = False
    
        'check if styles exist, else exit sub
        On Error GoTo ErrHandler:
        Dim keyStyle As Word.Style
    
        Set keyStyle = activeDoc.Styles(findStyle)
        Set keyStyle = activeDoc.Styles(prevStyle)
    
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
            .Style = activeDoc.Styles(findStyle)
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
    
    Do While Selection.Find.Execute = True And jCount < 200            'jCount so we don't get an infinite loop
        jCount = jCount + 1
        
        'Get number of current pagaraph, because we get an error if try to select before 1st para
        Dim rParagraphs As Range
        Dim CurPos As Long
         
        Selection.Range.Select  'select current ran
        CurPos = activeDoc.Bookmarks("\startOfSel").Start
        Set rParagraphs = activeDoc.Range(Start:=0, End:=CurPos)
        intCurrentPara = rParagraphs.Paragraphs.Count
        
        'DebugPrint intCurrentPara
        
        If intCurrentPara > 1 Then
            'select preceding paragraph
            Selection.Previous(Unit:=wdParagraph, Count:=1).Select
            pageNum = Selection.Information(wdActiveEndPageNumber)
        
                'Check if preceding paragraph style is correct
                If Selection.Style <> prevStyle Then
                    jString = jString & "** ERROR: Missing or incorrect " & prevStyle & " style before " _
                        & findStyle & " style on page " & pageNum & "." & vbNewLine & vbNewLine
                End If
            
                'If you're searching for a page break before, also check if manual page break is in paragraph
                If prevStyle = "Page Break (pb)" Then
                    If InStr(Selection.Text, Chr(12)) = 0 Then
                        jString = jString & "** ERROR: Missing manual page break on page " & pageNum & "." _
                            & vbNewLine & vbNewLine
                    End If
                End If
            
                'DebugPrint jString
        
            'move the selection back to original paragraph, so it won't be
            'selected again on next search
            Selection.Next(Unit:=wdParagraph, Count:=1).Select
        End If
        
    Loop
    
    'DebugPrint jString
    
    CheckPrevStyle = jString
    
    Exit Function
    
ErrHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then       'style doesn't exist in document
        Exit Function
    End If
End Function



Private Function BadTorStyles(ProgressBar2 As ProgressBar, StatusBar As String, _
  ProgressTitle As String, CurrentStories As Collection) As String
    'Called from GoodBadStyles sub if torDOTcom parameter is set to True.
    
    Dim paraStyle As String
    Dim activeParaCount As Integer
    
    Dim strCsvFileName As String
    Dim strLogInfo() As Variant
    ReDim strLogInfo(1 To 3)
    Dim strFullPathToCsv As String
    Dim arrTorStyles() As Variant
    Dim strLogDir As String
    Dim strPathToLogFile As String
    
    Dim intBadCount As Integer
    Dim activeParaRange As Range
    Dim pageNumber As Integer
    
    Dim N As Integer
    Dim M As Integer
    Dim strBadStyles As String
    Dim A As Long
    
    Dim TheOS As String
    TheOS = System.OperatingSystem
    Dim sglPercentComplete As Single
    Dim strStatus As String
    Dim varStory As Variant
    Dim currentStory As WdStoryType
    
    Application.ScreenUpdating = False
    
    ' This is the file we want to download
    strCsvFileName = "Styles_Bookmaker.csv"
    
    'List of styles approved for use in Bookmaker
    'Organized by approximate frequency in manuscripts (most freq at top)
    arrTorStyles = SharedFileInstaller.DownloadCSV(FileName:=strCsvFileName)
    
    activeParaCount = activeDoc.Paragraphs.Count
    
    For N = 1 To activeParaCount
 
        If N Mod 100 = 0 Then
            'Percent complete and status for progress bar (PC) and status bar (Mac)
            sglPercentComplete = (((N / activeParaCount) * 0.1) + 0.76)
            strStatus = "* Checking paragraph " & N & " of " & activeParaCount & " for approved Bookmaker styles..." & vbCr & StatusBar
    
            Call UpdateBarAndWait(Bar:=ProgressBar2, Status:=strStatus, Percent:=sglPercentComplete)
        End If
        
        For Each varStory In CurrentStories
          currentStory = varStory
            If N <= activeDoc.StoryRanges(currentStory).Paragraphs.Count Then
                paraStyle = activeDoc.StoryRanges(currentStory).Paragraphs(N).Style
                'DebugPrint paraStyle
                
                If Right(paraStyle, 1) = ")" Then
                    'DebugPrint "Current paragraph is: " & paraStyle
                    On Error GoTo ErrHandler
                    
                    intBadCount = -1        ' -1 because the array is base 0
                    
                    For M = LBound(arrTorStyles()) To UBound(arrTorStyles())
                        'DebugPrint arrTorStyles(M, 0)
                        
                        If paraStyle <> arrTorStyles(M, 0) Then
                            intBadCount = intBadCount + 1
                        Else
                            Exit For
                        End If
                    Next M
                    
                    'DebugPrint intBadCount
                    If intBadCount = UBound(arrTorStyles()) Then
                        Set activeParaRange = activeDoc.StoryRanges(currentStory).Paragraphs(N).Range
                        pageNumber = activeParaRange.Information(wdActiveEndPageNumber)
                        strBadStyles = strBadStyles & "** ERROR: Non-Bookmaker style on page " & pageNumber _
                            & " (Paragraph " & N & "):  " & paraStyle & vbNewLine & vbNewLine
                            'DebugPrint strBadStyles
                    End If
                
                End If
            End If
        Next A
ErrResume:
    
    Next N
    
    StatusBar = "* Checking paragraphs for approved Bookmaker styles..." & vbCr & StatusBar
    
    'DebugPrint strBadStyles
    
    BadTorStyles = strBadStyles
    Exit Function

ErrHandler:
    DebugPrint Err.Number & " " & Err.Description & " | " & Err.HelpContext
    If Err.Number = 5941 Or Err.Number = 5834 Then       'style is not in document
        Resume ErrResume
    End If

End Function



Private Function GetMetadata() As String
    Dim styleNameB(3) As String         ' must declare number of items in array here
    Dim bString(3) As String            ' and here
    Dim B As Integer
    Dim strTitleData As String
    
    Application.ScreenUpdating = False
    
    styleNameB(1) = "Titlepage Book Title (tit)"
    styleNameB(2) = "Titlepage Author Name (au)"
    styleNameB(3) = "span ISBN (isbn)"
    
    For B = 1 To UBound(styleNameB())
        bString(B) = MacroHelpers.GetText(styleNameB(B))
        If bString(B) <> vbNullString Then
            bString(B) = "** " & styleNameB(B) & " **" & vbNewLine & _
                        bString(B) & vbNewLine
        End If
        
        strTitleData = strTitleData & bString(B)
        
    Next B
                
    'DebugPrint strTitleData
    
    GetMetadata = strTitleData

End Function

Private Function IllustrationsList() As String
    Dim cString(1000) As String             'Max number of illustrations. Could be lower than 1000.
    Dim cCount As Integer
    Dim pageNumberC As Integer
    Dim strFullList As String
    Dim N As Integer
    Dim strSearchStyle As String
    
    Application.ScreenUpdating = False
    
    strSearchStyle = "Illustration holder (ill)"
    cCount = 0
    
    'Move selection to start of document
    Selection.HomeKey Unit:=wdStory
        
        ' Check if search style exists in document
        On Error GoTo ErrHandler
        Dim keyStyle As Style
        
        Set keyStyle = activeDoc.Styles(strSearchStyle)
    
        Selection.Find.ClearFormatting
        With Selection.Find
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .Style = activeDoc.Styles(strSearchStyle)
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
        
        cString(cCount) = "Page " & pageNumberC & ": " & Selection.Text & vbNewLine
        
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
    
    'DebugPrint strFullList
    
    IllustrationsList = strFullList
    
    Exit Function

ErrHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then
        IllustrationsList = ""
        Exit Function
    End If

End Function


Private Function CreateReportText(TemplateUsed As Boolean, errorList As String, metadata As String, illustrations As String, goodStyles As String) As String

    Application.ScreenUpdating = False
    
    Dim strReportText As String
        
    If TemplateUsed = False Then
        strReportText = strReportText & vbNewLine
        strReportText = strReportText & "------------------------STYLES IN USE--------------------------" & vbNewLine
        strReportText = strReportText & "It looks like you aren't using the Macmillan style template." & vbNewLine
        strReportText = strReportText & "That's OK, but if you would like more info about your document," & vbNewLine
        strReportText = strReportText & "just attach the Macmillan style template and apply the styles" & vbNewLine
        strReportText = strReportText & "throughout the document." & vbNewLine
        strReportText = strReportText & vbNewLine
        strReportText = strReportText & goodStyles
    Else
        If errorList = "" Then
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & "                 CONGRATULATIONS! YOU PASSED!" & vbNewLine
            strReportText = strReportText & " But you're not done yet. Please check the info listed below." & vbNewLine
            strReportText = strReportText & vbNewLine
        Else
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & "                             OOPS!" & vbNewLine
            strReportText = strReportText & "     Problems were found with the styles in your document." & vbNewLine
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & "---------------------------- ERRORS ---------------------------" & vbNewLine
            strReportText = strReportText & errorList
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & vbNewLine
        End If
            strReportText = strReportText & "--------------------------- METADATA --------------------------" & vbNewLine
            strReportText = strReportText & "If any of the information below is wrong, please fix the" & vbNewLine
            strReportText = strReportText & "associated styles in the manuscript." & vbNewLine
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & metadata
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & "----------------------- ILLUSTRATION LIST ---------------------" & vbNewLine
        
            If illustrations <> "no illustrations detected" & vbNewLine Then
                strReportText = strReportText & "Verify that this list of illustrations includes only the file" & vbNewLine
                strReportText = strReportText & "names of your illustrations." & vbNewLine
                strReportText = strReportText & vbNewLine
            End If
        
            strReportText = strReportText & illustrations
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & vbNewLine
            strReportText = strReportText & "-------------------- MACMILLAN STYLES IN USE ------------------" & vbNewLine
            strReportText = strReportText & goodStyles
    End If

    CreateReportText = strReportText
    
End Function

Private Function StylesInUse(ProgressBar As ProgressBar, Status As String, ProgTitle As String) As String
    'Creates a list of all styles in use, not just Macmillan styles
    'No list of bad styles
    'For use when no Macmillan template is attached
    
    Dim TheOS As String
    TheOS = System.OperatingSystem
    Dim sglPercentComplete As Single
    Dim strStatus As String
    
    Dim stylesGood() As String
    Dim stylesGoodLong As Long
    stylesGoodLong = 400                                    'could maybe reduce this number
    ReDim stylesGood(stylesGoodLong)
    Dim styleGoodCount As Integer
    Dim activeParaCount As Integer
    Dim J As Integer, K As Integer, L As Integer
    Dim paraStyle As String
    '''''''''''''''''''''
    Dim activeParaRange As Range
    Dim pageNumber As Integer
    
    Dim colStories As Collection
    Set colStories = MacroHelpers.ActiveStories
    Dim varStory As Variant
    Dim currentStory As WdStoryType
    
    '----------Collect all styles being used-------------------------------
    styleGoodCount = 0
    activeParaCount = activeDoc.Paragraphs.Count
    For J = 1 To activeParaCount
        
        'All Progress Bar statements for PC only because won't run modeless on Mac
        If J Mod 100 = 0 Then
        
            'Percent complete and status for progress bar (PC) and status bar (Mac)
            sglPercentComplete = (((J / activeParaCount) * 0.12) + 0.86)
            strStatus = "* Checking paragraph " & J & " of " & activeParaCount & " for Macmillan styles..." & vbCr & Status
    
            Call UpdateBarAndWait(Bar:=ProgressBar, Status:=strStatus, Percent:=sglPercentComplete)
            
        End If
        
        For Each varStory In colStories
          currentStory = varStory
            If J <= activeDoc.StoryRanges(currentStory).Paragraphs.Count Then
                paraStyle = activeDoc.StoryRanges(currentStory).Paragraphs(J).Style
                Set activeParaRange = activeDoc.StoryRanges(currentStory).Paragraphs(J).Range
                pageNumber = activeParaRange.Information(wdActiveEndPageNumber)                 'alt: (wdActiveEndAdjustedPageNumber)
        
                For K = 1 To styleGoodCount
                    ' "Left" function because now stylesGood includes page number, so won't match paraStyle
                    If paraStyle = Left(stylesGood(K), InStrRev(stylesGood(K), " --") - 1) Then
                        K = styleGoodCount                              'stylereport bug fix #1    v. 3.1
                        Exit For                                        'stylereport bug fix #1   v. 3.1
                    End If                                              'stylereport bug fix #1   v. 3.1
                Next K
                If K = styleGoodCount + 1 Then
                    styleGoodCount = K
                    stylesGood(styleGoodCount) = paraStyle & " -- p. " & pageNumber
                End If
            End If
        Next A
    Next J
    
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
    
    'DebugPrint strGoodStyles
    
    StylesInUse = strGoodStyles

End Function

Private Sub ISBNcleanup()
'removes "span ISBN (isbn)" style from all but the actual ISBN numerals
    
    'check if that style exists, if not then exit sub
    On Error GoTo ErrHandler:
        Dim keyStyle As Word.Style
        Set keyStyle = activeDoc.Styles("span ISBN (isbn)")
    On Error GoTo 0
    
    Dim strISBNtextArray()
    ReDim strISBNtextArray(1 To 3)
    
    strISBNtextArray(1) = "-[!0-9]"     'any hyphen followed by any non-digit character
    strISBNtextArray(2) = "[!0-9]-"     'any hyphen preceded by any non-digit character
    strISBNtextArray(3) = "[!-0-9]"     'any character other than a hyphen or digit
    
    ' re: above--need to search for hyphens first, because if you lead with what is now 3, you
    ' remove the style from any characters around hyphens, so if you search for a hyphen next to
    ' a character later, it won't return anything because the whole string needs to have the
    ' style applied for it to be found.
    
    Dim g As Long
    For g = LBound(strISBNtextArray()) To UBound(strISBNtextArray())
        
        'Move selection to start of document
        Selection.HomeKey Unit:=wdStory

        With Selection.Find
            .ClearFormatting
            .Text = strISBNtextArray(g)
            .Replacement.ClearFormatting
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .Style = "span ISBN (isbn)"                     'find this style
            .Replacement.Style = "Default Paragraph Font"   'replace with this style
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        
        Selection.Find.Execute Replace:=wdReplaceAll
    
    Next g
    
Exit Sub
    
ErrHandler:
    If Err.Number = 5941 Or Err.Number = 5834 Then       'Style doesn't exist in document
        Exit Sub
    End If

End Sub

Private Function BookTypeCheck()
    ' Validates the book types listed following the ISBN on the copyright page.
    Dim intCount As Integer
    Dim strErrors As String
    Dim strBookTypes(1 To 7) As String
    Dim A As Long
    Dim blnMissing As Boolean
    Dim strIsbn As String
    
    strBookTypes(1) = "trade paperback"
    strBookTypes(2) = "hardcover"
    strBookTypes(3) = "e-book"
    strBookTypes(4) = "ebook"
    strBookTypes(5) = "print on demand"
    strBookTypes(6) = "print-on-demand"
    strBookTypes(7) = "mass market paperback"
    
    'Move selection back to start of document
    Selection.HomeKey Unit:=wdStory

    On Error GoTo ErrHandler
    
    intCount = 0
    With Selection.Find
        .ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Style = activeDoc.Styles("span ISBN (isbn)")
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        Do While .Execute(Forward:=True) = True And intCount < 100   ' < 100 to precent infinite loop
            intCount = intCount + 1
            strIsbn = Selection.Text
            'Record current selection because we need to return to it later
            activeDoc.Bookmarks.Add Name:="ISBN", Range:=Selection.Range
            
            Selection.Collapse Direction:=wdCollapseEnd
            Selection.EndOf Unit:=wdLine, Extend:=wdExtend
            
            blnMissing = True
                For A = 1 To UBound(strBookTypes())
                    If InStr(Selection.Text, "(" & strBookTypes(A) & ")") > 0 Then
                        blnMissing = False
                        Exit For
                    End If
                Next A
            
            If blnMissing = True Then
                strErrors = strErrors & "** ERROR: Correct book type required in parentheses after" & vbNewLine & _
                    "ISBN " & strIsbn & " on copyright page." _
                    & vbNewLine & vbNewLine
            End If
            
            'Now we need to return the selection to where it was above, or else we can't loop through selection.find
            If activeDoc.Bookmarks.Exists("ISBN") = True Then
                Selection.GoTo what:=wdGoToBookmark, Name:="ISBN"
                activeDoc.Bookmarks("ISBN").Delete
            End If
            
        Loop
    
    End With
    
    'DebugPrint strErrors
    BookTypeCheck = strErrors
    
    On Error GoTo 0
    Exit Function

ErrHandler:
    DebugPrint Err.Number & ": " & Err.Description
    If Err.Number = 5941 Or Err.Number = 5834 Then      ' style doesn't exist in document
        Exit Function
    End If
        
End Function


Private Sub ChapNumCleanUp()
  On Error GoTo ErrHandler
  Dim strCNStyle As String
  strCNStyle = "Chap Number (cn)"

  If MacroHelpers.IsStyleInUse(strCNStyle) = True Then
  
    ' Removes character styles from Chapter Number paragraphs
    Dim iCount As Long
    Dim strText As String
    Dim intCount As Long

    'Move selection back to start of document
    Selection.HomeKey Unit:=wdStory
    
    intCount = 0
    With Selection.Find
        .ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .Style = activeDoc.Styles(strCNStyle)
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        Do While .Execute(Forward:=True) = True And intCount < 1000   ' < 1000 to prevent infinite loop
            intCount = intCount + 1
            #If Mac Then
                ' Mac 2011 doesn't support ClearCharacterFormattingAll method
                ' And ClearFormatting removes paragraph formatting as well
                Selection.ClearFormatting
                Selection.Style = strCNStyle
            #Else
                Selection.ClearCharacterAllFormatting
            #End If
        Loop
    
    End With
  End If
    
  Exit Sub
    
ErrHandler:
        'DebugPrint Err.Number & ": " & Err.Description
    If Err.Number = 5941 Or Err.Number = 5834 Then      ' style doesn't exist in document
        Exit Sub
    End If
End Sub


Private Function CheckFileName() As Boolean
' Returns error message if file name contains special characters

    Dim strDocName As String
    Dim strCheckChar As String
    Dim strAllGoodChars As String
    Dim lngNameLength As Long
    Dim R As Long
    Dim strErrorString As String
    
    CheckFileName = False
    
    ' Only alphanumeric, underscore and hyphen allowed in Bkmkr names
    ' Will do vbTextCompare later for case insensitive search
    strAllGoodChars = "ABCDEFGHIJKLMNOPQRSTUVWZYX1234567890_-"
    
    ' Get file name w/o extension
    strDocName = activeDoc.Name
    strDocName = Left(strDocName, InStrRev(strDocName, ".") - 1)
    
    lngNameLength = Len(strDocName)
    
    ' Loop: pull each char in file name, check if it appears in good char
    ' list. If it doesn't appear, then it's bad! So return True
    ' Error is same whether there is 1 or 100 bad chars, so exit as soon as
    ' one is found.
    
    For R = 1 To lngNameLength
        strCheckChar = Mid(strDocName, R, 1)
        If InStr(1, strAllGoodChars, strCheckChar, vbTextCompare) = 0 Then
            CheckFileName = True
            Exit Function
        End If
    Next R

End Function

' ===== SectionStartRules =====================================================
' Moved JSON download, creation of rule dictionary and ProcessRules loop to
' SSRulesCollection itself.

Private Function SectionStartRules() As String

  Dim objNewSSruleCollection As SSRuleCollection
'  Dim lngRuleCount As Long
'  Dim strRuleName As String
'  Dim lngRulePriority As Long
'  Dim lngPriorityCount As Long
'  Dim lngPriorityCheck As Long

  ' create collection object (which creates a collection of Rule objects)
  Set objNewSSruleCollection = New SSRuleCollection
  SectionStartRules = objNewSSruleCollection.Validate

  ' Loop through Rules by "priority" values (set in SSRule.cls)
'  lngPriorityCount = 1
'  lngPriorityCheck = 1
'  Do Until lngPriorityCheck = 0
'    lngPriorityCheck = 0
'      For lngRuleCount = 1 To objNewSSruleCollection.Rules.Count
'        If objNewSSruleCollection.Rules(lngRuleCount).Priority = lngPriorityCount Then
'          Call ProcessRule(objNewSSruleCollection.Rules(lngRuleCount), objNewSSruleCollection.SectionLists)
'          lngPriorityCheck = lngPriorityCheck + 1
'        End If
'      Next
'    lngPriorityCount = lngPriorityCount + 1
'  Loop

End Function

