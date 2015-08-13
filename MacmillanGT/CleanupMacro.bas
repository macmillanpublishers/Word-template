Attribute VB_Name = "CleanupMacro"
' by Erica Warren - erica.warren@macmillan.com

' ======== PURPOSE ============================
' Performs standard typographic cleanup in current document

' ======== DEPENDENCIES =======================
' 1. Requires ProgressBar userform module

Option Explicit
Option Base 1

Dim activeRng As Range

Sub MacmillanManuscriptCleanup()

    ''''''''''''''''''''''''''''''''
    '''created by Matt Retzer  - matthew.retzer@macmillan.com
    ''''''''''''''''''''''''''''''
    'version 3.9: Adding progress bar

    'version 3.8.2: adding handling of track changes

    'version 3.8.1: fixing bug in character styles macro that was causing page breaks to drop out

    'version 3.8:
    'updated 2015-03-25 by Erica Warren
    ' Design Note can now contain blank characters
    ' new char styles "span symbols ital (symi)" and "span symbols bold (symb)" added to existing character styles to tag
    'changed way error checks verify that template is attached
    '

    'version 3.7
    'updated 2015-03-04, by Erica Warren
    ' split Cleanup macro into two macros: cleanup and character styles
    ' add new character styles (from v. 3.5) to new styles macro
    ' cleanup now removes space between ellipsis and double or single quote
    ' cleanup now removes blank paragraph at end of document
    ' cleanup now converts double periods to single periods
    ' cleanup now converts double commas to single commas

    '''''''''''''''''''''''''''''''''''
    ' version 3.6 : style updates only, no macro updates

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'version 3.5
    'updated Erica Warren 2015-02-18
    ''style report opens when complete

    '''''''''''''''''''''''''''''''''''''''''''''''''
    'version 3.4.3
    'last updated 2014-10-20 by Erica Warren
    ''' - moved StylesHyperlink sub after PreserveWhiteSpaceinBrkStylesA sub to prevent styled blank paragraphs from being removed

    '''''''''''''''''''''''''''''''''''''''
    'version 3.4.2: template style updates only, not macro updates

    ''''''''''''''''''''''''''''''''''''''
    'version 3.4.1
    'last updated 2014-10-08 by Erica Warren
    ''' - added new Column Break (cbr) style to preserve white space macro
    
    '''''''''''''''''''''''''''''''''''''''
    'version 3.4
    'last updated 2014-10-07 by Erica Warren
    ''' - removed Section Break (sbr) from RmNonWildcardItems sub
    ''' - added RemoveBookmarks sub
    ''' - added StyleHyperlinks sub, removed hyperlinks stuff from earlier version
    
    '''''''''''''''''''''''''''''''''''''''
    'version 3.3.1
    'last updated 2014-09-17 by Erica Warren
    ''' - fixed space break style names that were changed in template
    ''''''''''''''''''''''''''''''''''''''
    'version 3.3
    'last updated 2014-09-16 by Erica Warren
    ''' - added to RmWhiteSpaceB:
    '''     - remove space before closing parens, closing bracket, closing braces
    '''     - remove space after opening parens, opening bracket, opening braces, dollar sign
    ''' - added double space to preserve character style search/replace
    
    '''''''''''''''''''''''''''''''''
    'version 3.2
    'last updated 2014-09-12 by Erica Warren - erica.warren@macmillan.com
    ''' - changed double- and single- quotes replace to find only straight quotes
    ''' - added PreserveStyledPageBreaksA and PreserveStyledPageBreaksB, now required for correct InDesign import
    ''' - added PC_BestStylesView, Mac_BestStylesView, and StylesViewLaunch macros
    ''' - edited some msgBox text to make it a little more fun
    
    '''''''''''''''''''''''''''''''
    '''version 3.1
    '''last updated 07/08/14:
    ''' - split Localstyle replace into to private subs
    ''' - style report bug fix
    ''' - adding f/ replace nbs, nbh, oh: to wildcard f/r
    ''' - added completion message for Cleanup macro
    ''' - changed TagHyperlink sub to tagexistingchrlinks, added 6 more hyperlink stylesarstyles, including hype
    ''' - added a backtick on closing tags for preserved break styles, and a call to remove paras trailing these breaks
    ''' - adding ' endash ' turn to emdash as per EW
    ''' - added 'save in place' msgbox for Cleanup macro.
    ''' - fixed embedded filed code hyperlink bug, just giving them a leading space
    ''' - prepared tagging for 'combinatrions' of local styles
    ''' - combined highlight removal with local style find loop
    ''' - combined smart quotes with existing no wildcard sub, made array/loop setup for same
    ''' - changing default tags for local and char styles to be asymmetrical:  `X|tagged item|X`
    ''' - updated error check for incidental tags to match asymmetric tags
    ''' - added in 3 combo styles to LocalFind and LocalReplace
    ''' - added status bar updates
    ''' - added additional repair to embedded hyperlink, also related to leading spaces (`Q` tag)
    ''' - update version in Document properties
    '''''''''''''
    '''version 3.0
    '''last updated 6/10/14:
    ''' - added Style Report Macro Sub
    ''' - added srErrorCheck Function
    '''version2.1 - 5/27/14:
    ''' - added 7 styles for preserving white space,
    ''' - preserving superscript & subscript - converting to char styles.
    ''' - added prelim checks for protected documents, incidental pre-existing backtick tags
    ''' - consolidated all preliminary error checks into one function
    ''' - updating char styles to match new prefixes, in style replacements, hyperlink finds, and errorcheck1
    ''' - fixed field object hyperlink bug
    ''' - add find/replace for any extra hyperlink tags `H`
    ''' - removed .Forward = True from all Find/Replaces as it is redundant when wrap = Continue
    ''' - made all Subs Private except for the Main one
    
    '----------Timer Start-----------------------------
    'Dim StartTime As Double
    'Dim SecondsElapsed As Double
    
    'Remember time when macro starts
    '  StartTime = Timer

    '-----------run preliminary error checks------------
    Dim exitOnError As Boolean
    
    exitOnError = zz_errorChecks()      ''Doc is unsaved, protected, or uses backtick character?
    If exitOnError <> False Then
    Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    '------------check for endnotes and footnotes--------------------------
    Dim stStories() As Variant
    Dim a As Long
    
    stStories = StoryArray
    
    '------------record status of current status bar and then turn on-------
    Dim currentStatusBar As Boolean
    currentStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    
    '--------Progress Bar------------------------------
    'Percent complete and status for progress bar (PC) and status bar (Mac)
    'Requires ProgressBar custom UserForm and Class
    Dim sglPercentComplete As Single
    Dim strStatus As String
    Dim strTitle As String
    
    'First status shown will be randomly pulled from array, for funzies
    Dim funArray() As String
    ReDim funArray(1 To 10)      'Declare bounds of array here

    funArray(1) = "* Waving magic wand..."
    funArray(2) = "* Doing a little dance..."
    funArray(3) = "* Making plans for the weekend..."
    funArray(4) = "* Setting sail for the tropics..."
    funArray(5) = "* Making a cup of tea..."
    funArray(6) = "* Beep beep boop beep..."
    funArray(7) = "* Writing the next Great American Novel..."
    funArray(8) = "* My, don't you look nice today..."
    funArray(9) = "* Having a snack..."
    funArray(10) = "* Initiating launch sequence..."
    
    Dim x As Integer
    
    'Rnd returns random number between (0,1], rest of expression is to return an integer (1,10)
    Randomize           'Sets seed for Rnd below to value of system timer
    x = Int(UBound(funArray()) * Rnd()) + 1
    
    'Debug.Print x
    
    strTitle = "Macmillan Manuscript Cleanup Macro"
    sglPercentComplete = 0.05
    strStatus = funArray(x)

    'All Progress Bar statements for PC only because won't run modeless on Mac
    Dim TheOS As String
    TheOS = System.OperatingSystem
    
    If Not TheOS Like "*Mac*" Then
        Dim oProgressCleanup As ProgressBar
        Set oProgressCleanup = New ProgressBar
    
        oProgressCleanup.Title = strTitle
        oProgressCleanup.Show
    
        oProgressCleanup.Increment sglPercentComplete, strStatus
        Doze 50 'Wait 50 milliseconds for progress bar to update
    Else
        'Mac will just use status bar
        Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
        DoEvents
    End If
    
    '--------save the current cursor location in a bookmark---------------------------
    ActiveDocument.Bookmarks.Add Name:="OriginalInsertionPoint", Range:=Selection.Range
    
    '-----------Turn off track changes--------
    Dim currentTracking As Boolean
    currentTracking = ActiveDocument.TrackRevisions
    ActiveDocument.TrackRevisions = False
    
    
    '-----------Find/Replace with Wildcards = False--------------------------------
    Call zz_clearFind                          'Clear find object
    
    sglPercentComplete = 0.2
    strStatus = "* Fixing quotes, unicode, section breaks..." & vbCr & strStatus
    
    If Not TheOS Like "*Mac*" Then
        oProgressCleanup.Increment sglPercentComplete, strStatus
        Doze 50 'Wait 50 milliseconds for progress bar to update
    Else
        'Mac will just use status bar
        Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
        DoEvents
    End If
    
    Dim s As Long
    
    For s = 1 To UBound(stStories())
        Call RmNonWildcardItems(StoryType:=s)   'has to be alone b/c Match Wildcards has to be disabled: Smart Quotes, Unicode (ellipse), section break
    Next s
    
    Call zz_clearFind

    '-------------Tag characters styled "span preserve characters"-----------------
    sglPercentComplete = 0.4
    strStatus = "* Preserving styled whitespace characters..." & vbCr & strStatus
    
    If Not TheOS Like "*Mac*" Then
        oProgressCleanup.Increment sglPercentComplete, strStatus
        Doze 50 'Wait 50 milliseconds for progress bar to update
    Else
        'Mac will just use status bar
        Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
        DoEvents
    End If
    
    For s = 1 To UBound(stStories())
        Call PreserveStyledCharactersA(StoryType:=s)              ' EW added v. 3.2, tags styled page breaks, tabs
    Next s
    Call zz_clearFind
    
    '---------------Find/Replace for rest of the typographic errors----------------------
    sglPercentComplete = 0.6
    strStatus = "* Removing unstyled whitespace, fixing ellipses and dashes..." & vbCr & strStatus
    
    If Not TheOS Like "*Mac*" Then
        oProgressCleanup.Increment sglPercentComplete, strStatus
        Doze 50 'Wait 50 milliseconds for progress bar to update
    Else
        'Mac will just use status bar
        Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
        DoEvents
    End If
    
    For s = 1 To UBound(stStories())
        Call RmWhiteSpaceB(StoryType:=s)    'v. 3.7 does NOT remove manual page breaks or multiple paragraph returns
    Next s
    
    Call zz_clearFind
    
    '---------------Remove tags from "span preserve characters"-------------------------
    sglPercentComplete = 0.8
    strStatus = "* Cleaning up styled whitespace..." & vbCr & strStatus
    
    If Not TheOS Like "*Mac*" Then
        oProgressCleanup.Increment sglPercentComplete, strStatus
        Doze 50 'Wait 50 milliseconds for progress bar to update
    Else
        'Mac will just use status bar
        Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
        DoEvents
    End If
    
    For s = 1 To UBound(stStories())
        Call PreserveStyledCharactersB(StoryType:=s)    ' EW added v. 3.2, replaces character tags with actual character
    Next s
    
    Call zz_clearFind
    
    '-------------Go back to original insertion point and delete bookmark-----------------
    If ActiveDocument.Bookmarks.Exists("OriginalInsertionPoint") = True Then
        Selection.GoTo what:=wdGoToBookmark, Name:="OriginalInsertionPoint"
        ActiveDocument.Bookmarks("OriginalInsertionPoint").Delete
    End If
    
    '--------------Remove other bookmarks------------------------------------------------
    sglPercentComplete = 0.95
    strStatus = "* Removing bookmarks..." & vbCr & strStatus
    
    If Not TheOS Like "*Mac*" Then
        oProgressCleanup.Increment sglPercentComplete, strStatus
        Doze 50 'Wait 50 milliseconds for progress bar to update
    Else
        'Mac will just use status bar
        Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
        DoEvents
    End If
    
    Call RemoveBookmarks                    'this is in both Cleanup macro and ApplyCharStyles macro
    Call zz_clearFind
    
    '-----------------Restore original settings--------------------------------------
    sglPercentComplete = 1#
    strStatus = "* Finishing up..." & vbCr & strStatus
    
    If Not TheOS Like "*Mac*" Then
        oProgressCleanup.Increment sglPercentComplete, strStatus
        Doze 50 'Wait 50 milliseconds for progress bar to update
    Else
        'Mac will just use status bar
        Application.StatusBar = strTitle & " " & (100 * sglPercentComplete) & "% complete | " & strStatus
        DoEvents
    End If
    
    ActiveDocument.TrackRevisions = currentTracking         'Return track changes to the original setting
    Application.DisplayStatusBar = currentStatusBar
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    
    If Not TheOS Like "*Mac*" Then
        Unload oProgressCleanup
    End If
    
    MsgBox "Hurray, the Macmillan Cleanup macro has finished running! Your manuscript looks great!"                                 'v. 3.1 patch / request  v. 3.2 made a little more fun
    
    '----------------Timer End-----------------
    'Determine how many seconds code took to run
    '  SecondsElapsed = Round(Timer - StartTime, 2)
    
    'Notify user in seconds
    '  Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"
  
End Sub

Private Sub RemoveBookmarks()
    
    Dim bkm As Bookmark
    
    For Each bkm In ActiveDocument.Bookmarks
        bkm.Delete
    Next bkm
    
End Sub

Private Sub RmNonWildcardItems(StoryType As WdStoryType)                                             'v. 3.1 patch : redid this whole thing as an array, addedsmart quotes, wrap toggle var
    Set activeRng = ActiveDocument.StoryRanges(StoryType)
    
    Dim noWildTagArray(3) As String                                   ' number of items in array should be declared here
    Dim noWildReplaceArray(3) As String              ' number of items in array should be declared here
    Dim c As Long
    Dim wrapToggle As String
    
    wrapToggle = "wdFindContinue"
    Application.Options.AutoFormatAsYouTypeReplaceQuotes = True
    
    
    noWildTagArray(1) = "^u8230"
    noWildTagArray(2) = "^39"                       'v. 3.2: EW changed to straight single quote only
    noWildTagArray(3) = "^34"                       'v. 3.2: EW changed to straight double quote only
    
    noWildReplaceArray(1) = " . . . "
    noWildReplaceArray(2) = "'"
    noWildReplaceArray(3) = """"
    
    For c = 1 To UBound(noWildTagArray())
        If c = 3 Then wrapToggle = "wdFindStop"
        
        With activeRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = noWildTagArray(c)
            .Replacement.Text = noWildReplaceArray(c)
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
    Next

End Sub

Private Sub PreserveStyledCharactersA(StoryType As WdStoryType)
    ' added by EW v. 3.2
    ' replaces correctly styled characters with placeholder so they don't get removed
    Set activeRng = ActiveDocument.StoryRanges(StoryType)
    
    Dim preserveCharFindArray(3) As String  ' declare number of items in array
    Dim preserveCharReplaceArray(3) As String   'delcare number of items in array
    Dim preserveCharStyle As String
    Dim M As Long
    
    preserveCharStyle = "span preserve characters (pre)"
    
    On Error GoTo ErrHandler
    
        Dim keyStyle As Word.Style
        Set keyStyle = ActiveDocument.Styles(preserveCharStyle)
    
    preserveCharFindArray(1) = "^t" 'tabs
    preserveCharFindArray(2) = "  "  ' two spaces
    preserveCharFindArray(3) = "   "    'three spaces
    
    preserveCharReplaceArray(1) = "`E|"
    preserveCharReplaceArray(2) = "`G|"
    preserveCharReplaceArray(3) = "`J|"
    
    For M = 1 To UBound(preserveCharFindArray())
        With activeRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = preserveCharFindArray(M)
            .Replacement.Text = preserveCharReplaceArray(M)
            .Wrap = wdFindContinue
            .Format = True
            .Style = preserveCharStyle
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Next
    
    Exit Sub
    
ErrHandler:
    ' 5834: Item with specified name does not exist
    ' 5941: Member is not part of the collection (style doesn't exist)
    If Err.Number = 5834 Or Err.Number = 5941 Then
        Exit Sub
    End If
    
End Sub

Private Sub RmWhiteSpaceB(StoryType As WdStoryType)
    Set activeRng = ActiveDocument.StoryRanges(StoryType)

    Dim wsFindArray(33) As String              'number of items in array should be declared here
    Dim wsReplaceArray(33) As String       'and here
    Dim i As Long

    wsFindArray(1) = ".{4,}"             '4 or more consecutive periods, into proper 4 dot ellipse
    wsFindArray(2) = "..."                  '3 consecutive periods, into 3 dot ellipse
    wsFindArray(3) = "^s"                  'non-breaking space replaced with space                                 v. 3.1 patch
    wsFindArray(4) = "([! ]). . ."          'add leading space for ellipse if not present
    wsFindArray(5) = ". . .([! ])"          'add trailing space for ellipse if not present
    wsFindArray(6) = "^t{1,}"             'tabs replace with spaces
    wsFindArray(7) = "^l{1,}"               'manual line breaks replaced with hard return
    wsFindArray(8) = " {2,}"               '2 or more spaces replaced with single space
    wsFindArray(9) = "^13 "               'paragraph, space replaced with just paragraph
    wsFindArray(10) = " ^13"               'space, paragraph replaced with just paragraph
    wsFindArray(11) = "^-"                     'optional hyphen deleted                                                    v. 3.1 patch
    wsFindArray(12) = "^~"                      'non-breaking hyphen replaced with reg hyphen               v. 3.1 patch
    wsFindArray(13) = " ^= "                    'endash w/ spaces convert to emdash (no spaces)                                v. 3.1 patch
    wsFindArray(14) = "---"                   '3 hyphens to emdash
    wsFindArray(15) = "--"                   '2 hyphens to emdash                           v. 3.7 changed from en-dash to em-dash, per usual usage.
    wsFindArray(16) = " -"                  'hyphen leading space-remove
    wsFindArray(17) = "- "                  'hyphen trailing space-remove
    wsFindArray(18) = " ^+"                  'emdash leading space-remove
    wsFindArray(19) = "^+ "                  'emdash trailing space-remove
    wsFindArray(20) = " ^="                  'endash leading space-remove
    wsFindArray(21) = "^= "                  'endash trailing space-remove
    wsFindArray(22) = "\( "                     ' remove space after open parens                                                           v. 3.3
    wsFindArray(23) = " \)"                     ' removespace before closing parens                                                       v. 3.3
    wsFindArray(24) = "\[ "                     ' removespace after opening bracket                                                    v. 3.3
    wsFindArray(25) = " \]"                    ' removespace before closing bracket                                                   v. 3.3
    wsFindArray(26) = "\{ "                     ' removespace after opening curly bracket                                          v. 3.3
    wsFindArray(27) = " \}"                     ' removespace before closing curly bracket                                         v. 3.3
    wsFindArray(28) = "$ "                      ' removespace after dollar sign                                                                v. 3.3
    wsFindArray(29) = " . . . ."                ' remove space before 4-dot ellipsis (because it's a period)       v 3.7
    wsFindArray(30) = ".."                         'replace double period with single period                v. 3.7
    wsFindArray(31) = ",,"                          'replace double commas with single comma                v. 3.7

    'Test if Mac or PC because character code for closing quotes is different on different platforms            v 3.7
    #If Mac Then
        'I am a Mac and will test if it is Word 2011 or higher
        If Val(Application.Version) > 14 Then
            'remove space between ellipsis and closing double quote on Mac
            wsFindArray(32) = ". . . " & Chr(211)
        End If
    #Else
        'I am Windows
        ' remove space between ellipsis and closing double quote on Windows
        wsFindArray(32) = ". . . " & Chr(148)
    #End If
        
    #If Mac Then
        'I am a Mac and will test if it is Word 2011 or higher
        If Val(Application.Version) > 14 Then
            'remove space between ellipsis and closing single quote on Mac
            wsFindArray(33) = ". . . " & Chr(213)
        End If
    #Else
        'I am Windows
        ' remove space between ellipsis and closing single quote on Windows
        wsFindArray(33) = ". . . " & Chr(146)
    #End If

    wsReplaceArray(1) = ". . . . "      ' v. 3.2 EW removed leading space--not needed, 1st dot is a period
    wsReplaceArray(2) = " . . . "
    wsReplaceArray(3) = " "
    wsReplaceArray(4) = "\1 . . ."
    wsReplaceArray(5) = ". . . \1"
    wsReplaceArray(6) = " "
    wsReplaceArray(7) = "^p"
    wsReplaceArray(8) = " "
    wsReplaceArray(9) = "^p"
    wsReplaceArray(10) = "^p"
    wsReplaceArray(11) = ""
    wsReplaceArray(12) = "-"
    wsReplaceArray(13) = "^+"
    wsReplaceArray(14) = "^+"
    wsReplaceArray(15) = "^+"       'v. 3.7 changed to em-dash per common usage
    wsReplaceArray(16) = "-"
    wsReplaceArray(17) = "-"
    wsReplaceArray(18) = "^+"
    wsReplaceArray(19) = "^+"
    wsReplaceArray(20) = "^="
    wsReplaceArray(21) = "^="
    wsReplaceArray(22) = "("
    wsReplaceArray(23) = ")"
    wsReplaceArray(24) = "["
    wsReplaceArray(25) = "]"
    wsReplaceArray(26) = "{"
    wsReplaceArray(27) = "}"
    wsReplaceArray(28) = "$"
    wsReplaceArray(29) = ". . . ."
    wsReplaceArray(30) = "."
    wsReplaceArray(31) = ","

    #If Mac Then
        If Val(Application.Version) > 14 Then
            wsReplaceArray(32) = ". . ." & Chr(211)
        End If
    #Else
        wsReplaceArray(32) = ". . ." & Chr(148)
    #End If

    #If Mac Then
        If Val(Application.Version) > 14 Then
            wsReplaceArray(33) = ". . ." & Chr(213)
        End If
    #Else
        wsReplaceArray(33) = ". . ." & Chr(146)
    #End If

    For i = 1 To UBound(wsFindArray())
        With activeRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = wsFindArray(i)
            .Replacement.Text = wsReplaceArray(i)
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Next

End Sub

Private Sub PreserveStyledCharactersB(StoryType As WdStoryType)
    ' added by EW v. 3.2
    ' replaces placeholders with original characters
    Set activeRng = ActiveDocument.StoryRanges(StoryType)

    Dim preserveCharFindArray(3) As String  ' declare number of items in array
    Dim preserveCharReplaceArray(3) As String   'declare number of items in array
    Dim preserveCharStyle As String
    Dim N As Long

    preserveCharStyle = "span preserve characters (pre)"

    On Error GoTo ErrHandler
        Dim keyStyle As Word.Style
        Set keyStyle = ActiveDocument.Styles(preserveCharStyle)

    preserveCharFindArray(1) = "`E|" 'tabs
    preserveCharFindArray(2) = "`G|"    ' two spaces
    preserveCharFindArray(3) = "`J|"   'three spaces

    preserveCharReplaceArray(1) = "^t"
    preserveCharReplaceArray(2) = "  "
    preserveCharReplaceArray(3) = "   "

    For N = 1 To UBound(preserveCharFindArray())
        With activeRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = preserveCharFindArray(N)
            .Replacement.Text = preserveCharReplaceArray(N)
            .Wrap = wdFindContinue
            .Format = True
            .Style = preserveCharStyle
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Next

ErrHandler:
    ' 5834 = Item with specified name does not exist
    ' 5941 = Style missing from style collection
    If Err.Number = 5834 Or Err.Number = 5941 Then
        Exit Sub
    End If
End Sub

Function zz_errorChecks()

    zz_errorChecks = False
    Dim mainDoc As Document
    Set mainDoc = ActiveDocument
    Dim iReply As Integer

    '-----make sure document is saved
    Dim docSaved As Boolean                                                                                                 'v. 3.1 update
    docSaved = mainDoc.Saved
    
    If docSaved = False Then
        iReply = MsgBox("Your document '" & mainDoc & "' contains unsaved changes." & vbNewLine & vbNewLine & _
            "Click OK and I will save your document and run the macro." & vbNewLine & vbNewLine & "Click 'Cancel' to exit.", _
            vbOKCancel, "Alert")
        
        If iReply = vbOK Then
            mainDoc.Save
        Else
            zz_errorChecks = True
            Exit Function
        End If
    
    End If

    '-----test protection
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "Uh oh ... protection is enabled on document '" & mainDoc & "'." & vbNewLine & _
            "Please unprotect the document and run the macro again." & vbNewLine & vbNewLine & _
            "TIP: If you don't know the protection password, try pasting contents of this file into a " & _
            "new file, and run the macro on that.", , "Error 2"
        zz_errorChecks = True
        Exit Function
    End If

    '-----test if backtick style tag already exists
    Set activeRng = ActiveDocument.Range
    Application.ScreenUpdating = False

    Dim existingTagArray(3) As String   ' number of items in array should be declared here
    Dim b As Long
    Dim foundBad As Boolean
    foundBad = False

    existingTagArray(1) = "`[0-9]`"
    existingTagArray(2) = "`[A-Z]|"
    existingTagArray(3) = "|[A-Z]`"

    For b = 1 To UBound(existingTagArray())
        With activeRng.Find
            .ClearFormatting
            .Text = existingTagArray(b)
            .Wrap = wdFindContinue
            .MatchWildcards = True
        End With

        If activeRng.Find.Execute Then foundBad = True: Exit For
    Next

    If foundBad = True Then                'If activeRng.Find.Execute Then
        MsgBox "Something went wrong! The macro cannot be run on Document:" & vbNewLine & "'" & mainDoc & _
            "'" & vbNewLine & vbNewLine & "Please contact Digital Workflow group for support, I am sure they will " & _
            "be happy to help.", , "Error Code: 1"
        zz_errorChecks = True
    End If

End Function
