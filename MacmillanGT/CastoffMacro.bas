Attribute VB_Name = "CastoffMacro"
Option Explicit
Sub UniversalCastoff()
' created by Erica Warren - erica.warren@macmillan.com

' ========== PUROPOSE ========================
' Takes user inputs from userform to calculate a castoff (estimated print page count) based on
' the current text of the document.

' ========== DEPENDENCIES ====================
' 1. Requires SharedMacros module to be installed in same template
' 2. Requires design character count CSV files and spine size files be saved as attachments to
'    https://confluence.macmillan.com/display/PBL/Word+Template+downloads+-+production
' 3. Requires CastoffForm userform module and TextBoxEventHandler class module. Input validation is done there.


    '---------- Check if doc is saved ---------------------------------
    'If CheckSave = True Then
    '    Exit Sub
    'End If
    
    '----------Load userform------------------------
    Dim objCastoffForm As CastoffForm
    Set objCastoffForm = New CastoffForm
    
    objCastoffForm.Show
        
    'If user selected 'Oops, No Castoff For Me" button, cancel macro
    'If objCastoffForm.blnCancel = True Then
    '    Unload objCastoffForm
    '    Exit Sub
    'Else
    '    Call CastoffStart(FormInputs:=objCastoffForm)
    'End If
    
End Sub

Public Sub CastoffStart(FormInputs As CastoffForm)
    
    ' ============================================
    ' FOR TESTING / DEBUGGING
    ' If set to true, downloads CSV files from https://confluence.macmillan.com/display/PBL/Word+template+downloads+-+staging
    ' instead of production page (noted above)
    Dim blnStaging As Boolean
    blnStaging = False
    ' ============================================
    
    '----------Get user inputs from Userform--------------------------------------------------
    
    'Get Title Information
    Dim strEditor As String
    strEditor = FormInputs.txtEditor
    
    Dim strAuthor As String
    strAuthor = FormInputs.txtAuthor
    
    Dim strTitle As String
    strTitle = FormInputs.txtTitle
    
            
    'Get publisher name from option buttons
    Dim strPub As String        'Publisher code for file names and stuff
    Dim strPubRealName As String    'Publisher name for final output
    If FormInputs.optPubSMP.Enabled Then
        strPub = "SMP"
        strPubRealName = "St. Martin's Press"
    ElseIf FormInputs.optPubTor.Enabled Then
        strPub = "torDOTcom"
        strPubRealName = "Tor.com"
    ElseIf FormInputs.optPubPickup.Enabled Then
        ' Go to just calculating Pickup info
        ' ==== ADD CALL TO PICKUP SUB HERE =====
    End If
    
    
    'Get trim size
    Dim intTrim As Integer      'Index number of trim size in CSV file, starts at 0
    Dim strTrim As String       'Text version of trim size
    
    'Number assigned is column index in design array
    '0 = 5-1/2 x 8-1/4
    '1 = 6-1/8 x 9-1/4
    
    If FormInputs.optTrim5x8 Then
        intTrim = 0
        strTrim = "5-1/2 x 8-1/4"
    ElseIf FormInputs.optTrim6x9 Then
        intTrim = 1
        strTrim = "6-1/8 x 9-1/4"
    Else
        MsgBox "You must select a Trim Size to run the Castoff Macro." 'redundant since we're validating in the form module
        FormInputs.Show
    End If
        
    'Get designs selected.
    Dim intDesign() As Integer  'Index number of design density in CSV file, starts at 0
    Dim strDesign() As String   'Text of design density
    Dim intDim As Integer       'Number of dimensions  of intDesign and strDesign
    
    intDim = 0
    
    'For each design checked, increase dimension by 1 and then assign index and text of the design to an array
    If FormInputs.chkDesignLoose Then
        intDim = intDim + 1
        ReDim Preserve intDesign(1 To intDim)
        ReDim Preserve strDesign(1 To intDim)
        intDesign(intDim) = 0
        strDesign(intDim) = FormInputs.chkDesignLoose.Caption
    End If
    
    If FormInputs.chkDesignAverage Then
        intDim = intDim + 1
        ReDim Preserve intDesign(1 To intDim)
        ReDim Preserve strDesign(1 To intDim)
        intDesign(intDim) = 1
        strDesign(intDim) = FormInputs.chkDesignAverage.Caption
    End If
    
    If FormInputs.chkDesignTight Then
        intDim = intDim + 1
        ReDim Preserve intDesign(1 To intDim)
        ReDim Preserve strDesign(1 To intDim)
        intDesign(intDim) = 2
        strDesign(intDim) = FormInputs.chkDesignTight.Caption
    End If
    
    ' Get info from Standard Items section
    Dim intChapters As Integer      ' number of chapters
    intChapters = FormInputs.numTxtChapters
    
    Dim intParts As Integer         'number of part titles
    intParts = FormInputs.numTxtParts
    
    Dim intFMpgs As Integer         ' number of pages of frontmatter including blanks
    intFMpgs = FormInputs.numTxtFrontmatter
    
    ' Get info from Back Matter section
    Dim intIndexPgs As Integer     'Number of pages estimated for the index
    intIndexPgs = FormInputs.numTxtIndex.Text
    
    Dim intBackmatterPgsTK As Integer 'Number of pages of backmatter TK
    intBackmatterPgsTK = FormInputs.numTxtBackmatter.Text
    
    ' Get info from Notes and Bibliography section
    Dim intUnlinkedNotesPgs As Integer 'Number of unlinked endnotes in manuscript
    intUnlinkedNotesPgs = FormInputs.numTxtUnlinkedNotes.Text
    
    Dim intEndnotesPgsTK As Integer 'Number of endnotes pages TK
    intEndnotesPgsTK = FormInputs.numTxtNotesTK.Text
    
    Dim intBiblioPgs As Integer 'Number of pages of of bibliography currently in MS
    intBiblioPgs = FormInputs.numTxtBibliography.Text
    
    Dim intBiblioPgsTK As Integer 'Number of pages of Bibliography TK
    intBiblioPgsTK = FormInputs.numTxtBiblioTK.Text
    
    ' Get info from Complex Items section
    Dim intSubheads2Chap As Integer 'Number of subheads in 2 chapters
    intSubheads2Chap = FormInputs.numTxtSubheads.Text
    
    Dim intTablesPgs As Integer  'Number of pages for tables
    intTablesPgs = FormInputs.numTxtTables.Text
    
    Dim intArtPgs As Integer  'Number of pages for in-text art
    intArtPgs = FormInputs.numTxtArt.Text
    
    ' -------- Get character counts of current document -----------------
    
    
    
    ' -------- Do stuff with the userform inputs ------------------------
    
    '---------Download CSV with design specs from Confluence site-------
    Dim strCastoffFile As String    'File name of CSV on Confluence
    Dim strInfoType As String       'Type of into we want to download (here = Castoff)
    
    'Create name of castoff csv file to download
    strInfoType = "Castoff"
    strCastoffFile = strInfoType & "_" & strPub & ".csv"

    
    'Create log file name
    Dim arrLogInfo() As Variant
    ReDim arrLogInfo(1 To 3)
    
    arrLogInfo() = CreateLogFileInfo(strCastoffFile)
      
    'Create final path for downloaded CSV file (in log directory)
    'not in temp dir because that is where DownloadFromConfluence downloads it to, and it cleans that file up when done
    Dim strStyleDir As String
    Dim strPath As String
    Dim strLogFile As String
    Dim strMessage As String
    Dim strDir As String
    
    strStyleDir = arrLogInfo(1)
    strDir = arrLogInfo(2)
    strLogFile = arrLogInfo(3)
    strPath = strDir & Application.PathSeparator & strCastoffFile
        
    'Check if log file already exists; if not, create it
    CheckLog strStyleDir, strDir, strLogFile
    
    'Download CSV file from Confluence
    If DownloadFromConfluence(blnStaging, strDir, strLogFile, strCastoffFile) = False Then
        ' If download fails, check if we have an older version of the design CSV to work with
        If IsItThere(strPath) = False Then
            strMessage = "Looks like we can't download the design info from the internet right now. " & _
                "Please check your internet connection, or contact workflows@macmillan.com."
            MsgBox strMessage, vbCritical, "Error 5: Download failed, no previous design file available"
            Unload FormInputs
            Exit Sub
        Else
            strMessage = "Looks like we can't download the most up to date design info from the internet right now, " & _
                "so we'll just use the info we have on file for your castoff."
            MsgBox strMessage, vbInformation, "Let's do this thing!"
        End If
    End If
             
    'Double check that CSV is there
    If IsItThere(strPath) = False Then
        strMessage = "The Castoff macro is unable to access the design count file right now. Please check your internet " & _
                    "connection and try again, or contact workflows@macmillan.com."
        MsgBox strMessage, vbCritical, "Error 3: Design CSV doesn't exist"
        Unload FormInputs
        Exit Sub
    Else
        ' Load CSV into an array
        Dim arrDesign() As Variant
        arrDesign = LoadCSVtoArray(strPath)
    End If
    

            
    '------------Get castoff for each Design selected-------------------
    Dim lngDesignCount As Long
    Dim strWarning As String
    Dim strSpineSize As String
    Dim d As Long
    
    strWarning = ""
    strSpineSize = ""
    
    For d = LBound(intDesign()) To UBound(intDesign())
        'Debug.Print _
        UBound(arrDesign(), 1) & " < " & intDesign(d) & vbNewLine & _
        UBound(arrDesign(), 2) & " < "; intTrim
        
        'Error handling: intDesign(d) must be in range of design array
        If UBound(arrDesign(), 1) < intDesign(d) And UBound(arrDesign(), 2) < intTrim Then
             MsgBox "There was an error generating your castoff. Please contact workflows@macmillan.com for assistance.", _
                vbCritical, "Error 1: Design Count Out of Range"
            Unload FormInputs
            Exit Sub
        Else
            '---------Get design character count-------------------------------
            lngDesignCount = arrDesign(intDesign(d), intTrim)
    
            '--------------------------------------------------
            'For Reference: Index numbers in array (base 0)
            '
            '       | 5-1/2 x 8-1/4 |  6-1/8 x 9-1/4
            'loose  | (0,0)         | (0,1)
            'average| (1,0)         | (1,1)
            'tight  | (2,0)         | (2,1)
            '--------------------------------------------------

            'Debug.Print lngDesignCount

            '---------Calculate Page Count--------------------------------------
            Dim arrCastoffResult() As Variant
            Dim lngFinalCount As Long
            Dim lngBlankPgs As Long
            Dim lngActualCount As Long

            arrCastoffResult = Castoff(lngDesignCount, FormInputs)

            lngFinalCount = arrCastoffResult(0)
            lngBlankPgs = arrCastoffResult(1)
            lngActualCount = arrCastoffResult(2)

            'Add extra space if blanks less than 10
            Dim strExtraSpace As String
    
            If lngBlankPgs < 10 Then
                strExtraSpace = "    "
            Else
                strExtraSpace = "  "
            End If
    
            '---------Tor.com POD exceptions---------------------------------
            If strPub = "torDOTcom" Then
        
                'POD only has to be even, not 16-page sig
                If (lngActualCount Mod 2) = 0 Then      'page count is even
                    lngFinalCount = lngActualCount
                    lngBlankPgs = 0
                Else                                    'page count is odd
                    lngFinalCount = lngActualCount + 1
                    lngBlankPgs = 1
                End If
        
                'Warning about sub 48 page saddle-stitched tor.com books, warn if close to that
                If lngFinalCount < 56 Then
                    strWarning = "NOTE: Tor.com titles less than 48 pages will be saddle-stitched." & _
                                    vbNewLine & vbNewLine
                End If
            
                'Debug.Print strPub
                
                'Get spine size
                If lngFinalCount >= 18 And lngFinalCount <= 1050 Then       'Limits of spine size table
                    strSpineSize = SpineSize(blnStaging, lngFinalCount, strPub, FormInputs, strLogFile)
                    'Debug.Print "spine size = " & strSpineSize
                    If strSpineSize = vbNullString Then
                        strSpineSize = "Error 2: Word was unable to generate a spine size. " & _
                            "Contact workflows@macmillan.com for assistance."
                    Else
                        strSpineSize = "Your spine size will be " & strSpineSize & " inches " & _
                                            "at this page count."
                    End If
                Else
                    strSpineSize = "Your page count of " & lngFinalCount & _
                            " is out of range of the spine-size table."
                End If
    
            End If
            
            '------------------Create output for this castoff---------------------
            strMessage = strMessage & _
                vbTab & UCase(strDesign(d)) & ": " & lngFinalCount & vbNewLine & _
                vbTab & lngActualCount & " text pages" & vbNewLine & _
                vbTab & strExtraSpace & lngBlankPgs & " blank pages" & vbNewLine & _
                vbTab & lngFinalCount & " total pages" & vbNewLine & vbNewLine
        End If
    
    Next d
    
    '-------------Create final message---------------------------------------------------
    If strMessage <> vbNullString Then
        strMessage = "Your " & strPub & " title will have these approximate page counts" _
            & vbNewLine & "at the " & strTrim & " trim size:" & vbNewLine & vbNewLine & _
            strMessage
    Else
        strMessage = "Error 4: There was a problem generating your castoff." & _
            " Please contact workflows@macmillan.com for assistance."
    End If

    '-------------Report castoff info to user----------------------------------------------------------------
    MsgBox strMessage & strWarning & strSpineSize, vbOKOnly, "Castoff"

    Unload FormInputs
            
End Sub

Private Function LoadCSVtoArray(Path As String) As Variant

'------Load CSV into 2d array, NOTE!!: base 0---------
    Dim fnum As Integer
    Dim whole_file As String
    Dim lines As Variant
    Dim one_line As Variant
    Dim num_rows As Long
    Dim num_cols As Long
    Dim the_array() As Variant
    Dim R As Long
    Dim c As Long
    
        If IsItThere(Path) = False Then
            MsgBox "There was a problem with your Castoff.", vbCritical, "Error"
            Exit Function
        End If
        'Debug.Print Path

        ' Load the csv file.
        fnum = FreeFile
        Open Path For Input As fnum
        whole_file = Input$(LOF(fnum), #fnum)
        Close fnum

        ' Break the file into lines.
        lines = Split(whole_file, vbCrLf)

        ' Dimension the array.
        num_rows = UBound(lines)
        one_line = Split(lines(0), ",")
        num_cols = UBound(one_line)
        ReDim the_array(num_rows, num_cols)

        ' Copy the data into the array.
        For R = 0 To num_rows
            If Len(lines(R)) > 0 Then
                one_line = Split(lines(R), ",")
                For c = 0 To num_cols
                    the_array(R, c) = one_line(c)
                Next c
            End If
        Next R
    
        ' Prove we have the data loaded.

        ' For R = 0 To num_rows
        '     For c = 0 To num_cols
        '         Debug.Print the_array(R, c) & "|";
        '     Next c
        '     Debug.Print
        ' Next R
        ' Debug.Print "======="
        
        ' Delete the .csv file (actually keep it in case we need it later!)
        ' If Len(Dir$(Path)) > 0 Then
        '     Kill Path
        ' End If
    
    LoadCSVtoArray = the_array
    
End Function

Private Function Castoff(Design As Long, objForm As CastoffForm) As Variant
    Dim lngCharacterCount As Long
    Dim lngActualPageCount As Long
    Dim lngFinalPageCount As Long

    'Get character count with space from Word doc, divide by avg. char. count of design to get page count
    lngCharacterCount = ActiveDocument.Range.ComputeStatistics(wdStatisticCharactersWithSpaces)
    lngActualPageCount = lngCharacterCount / Design

    'Debug.Print "Starting page count: " & lngActualPageCount
    
    lngFinalPageCount = lngActualPageCount

    'search for page breaks, add a page for each
    'to account for space at end of chapters / beginning of chapters.
    'Editors should add blank pages for part title verso, etc.
    'Page breaks should be added as manual page breaks, not paragraph returns
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
        lngFinalPageCount = lngFinalPageCount + 1
    Loop
    End With

    'Debug.Print "Page count with page breaks: " & lngFinalPageCount
    
    'Add any missing pages indicated by user
    lngFinalPageCount = lngFinalPageCount + objForm.txtMissingPages.Text
    
    'Debug.Print "Page count with missing added: " & lngFinalPageCount
    
    'Determine next 16-page signature and blank pages
    Dim lngBlanks As Long
    Dim lngPgsOver As Long
    lngBlanks = 0
    lngPgsOver = (lngFinalPageCount Mod 16)

    If lngPgsOver <> 0 Then
        lngBlanks = 16 - lngPgsOver
        lngFinalPageCount = lngFinalPageCount + lngBlanks
        
    Else
        lngBlanks = 0
    End If

    'Debug.Print "Final page count: " & lngFinalPageCount
    'Debug.Print "Final blank pages: " & lngBlanks


    Dim arrResult() As Variant
    ReDim arrResult(0 To 2)

    arrResult(0) = lngFinalPageCount
    arrResult(1) = lngBlanks
    arrResult(2) = lngFinalPageCount - lngBlanks

    Castoff = arrResult

End Function
Private Function SpineSize(Staging As Boolean, PageCount As Long, Publisher As String, objForm As CastoffForm, LogFile As String)

'----Get Log dir to save spines CSV to --------------------------
    Dim strLogDir As String
    strLogDir = Left(LogFile, InStrRev(LogFile, Application.PathSeparator) - 1)
    'Debug.Print strLogDir

'----Define spine chart file name--------------------------------
    Dim strSpineFile As String
    strSpineFile = "Spine_" & Publisher & ".csv"
    
'----Define full path to where CSV will be----------------------
    Dim strFullPath As String
    strFullPath = strLogDir & Application.PathSeparator & strSpineFile
    
'----Download CSV with spine sizes from Confluence site----------
    Dim strMessage As String
    
    'Check if log file already exists; if not, create it then download CSV file
    If IsItThere(LogFile) = True Then
        If DownloadFromConfluence(Staging, strLogDir, LogFile, strSpineFile) = False Then
            ' If download fails, check if we have an older version of the spine CSV to work with
            If IsItThere(strFullPath) = False Then
                strMessage = "Looks like we can't download the spine size info from the internet right now. " & _
                    "Please check your internet connection, or contact workflows@macmillan.com."
                MsgBox strMessage, vbCritical, "Error 5: Download failed, no previous spine file available"
                Exit Function
            Else
                strMessage = "Looks like we can't download the most up to date spine size info from the internet right now, " & _
                    "so we'll just use the info we have on file for your castoff."
                MsgBox strMessage, vbInformation, "Let's do this thing!"
            End If
        End If
    End If
                        
    'Make sure CSV is there
    If IsItThere(strFullPath) = False Then
        strMessage = "The Castoff macro is unable to access the spine size file right now. Please check your internet " & _
                    "connection and try again, or contact workflows@macmillan.com."
        MsgBox strMessage, vbCritical, "Error 4: Spine CSV doesn't exist"
        Exit Function
    Else
        ' Load CSV into an array
        Dim arrDesign() As Variant
        arrDesign = LoadCSVtoArray(strFullPath)
    End If



    
'---------Lookup spine size in array-------------------------------
    Dim strSpine As String
    Dim c As Long
    
    For c = LBound(arrDesign, 1) To UBound(arrDesign, 1)
        'Debug.Print arrDesign(c, 0) & " = " & PageCount
        If arrDesign(c, 0) = PageCount Then
            strSpine = arrDesign(c, 1)
            Exit For
        End If
    Next c
    
    'Debug.Print strSpine
    SpineSize = strSpine

End Function
