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
    If CheckSave = True Then
        Exit Sub
    End If
    
    '----------Load userform------------------------
    Dim objCastoffForm As CastoffForm
    Set objCastoffForm = New CastoffForm
    
    objCastoffForm.Show
    
End Sub

Public Sub CastoffStart(FormInputs As CastoffForm)
    
    ' ============================================
    ' FOR TESTING / DEBUGGING
    ' If set to true, downloads CSV files from https://confluence.macmillan.com/display/PBL/Word+template+downloads+-+staging
    ' instead of production page (noted above)
    Dim blnStaging As Boolean
    blnStaging = False
    ' ============================================
    

    ' Get publisher name from option buttons
    Dim strPub As String            ' Publisher code for file names and stuff
    Dim strPubRealName As String    ' Publisher name for final output
    If FormInputs.optPubSMP Then
        strPub = "SMP"              ' CSV file on Confluence must use this code
        strPubRealName = "St. Martin's Press"
    ElseIf FormInputs.optPubTor Then
        strPub = "torDOTcom"        ' CSV file on Confluence must use this code
        strPubRealName = "Tor.com"
    ElseIf FormInputs.optPubPickup Then
        strPub = "Pickup"
        strPubRealName = "Pickup Design"
    End If
        
    'Debug.Print strPubRealName
    ' Get estimated page count
    Dim lngCastoffResult() As Long
    
    If strPub = "Pickup" Then
        ' If it's a pickup, use the PickupDesign math to get the page count, only 1 result
        ReDim Preserve lngCastoffResult(0)
        lngCastoffResult(0) = PickupDesign(FormInputs)
    Else
        ' Get designs selected from Form.
        Dim lngDesign() As Long     ' Index number of design density in CSV file, starts at 0
        Dim strDesign() As String   ' Text of design density
        Dim lngDim As Long       ' Number of dimensions of lngDesign and strDesign, base 1
        
        lngDim = -1
        
        'For each design checked, increase dimension by 1 and then assign index and text of the design to an array
        If FormInputs.chkDesignLoose Then
            lngDim = lngDim + 1
            ReDim Preserve lngDesign(0 To lngDim)
            ReDim Preserve strDesign(0 To lngDim)
            lngDesign(lngDim) = 0
            strDesign(lngDim) = FormInputs.chkDesignLoose.Caption
        End If
        
        If FormInputs.chkDesignAverage Then
            lngDim = lngDim + 1
            ReDim Preserve lngDesign(0 To lngDim)
            ReDim Preserve strDesign(0 To lngDim)
            lngDesign(lngDim) = 1
            strDesign(lngDim) = FormInputs.chkDesignAverage.Caption
        End If
        
        If FormInputs.chkDesignTight Then
            lngDim = lngDim + 1
            ReDim Preserve lngDesign(0 To lngDim)
            ReDim Preserve strDesign(0 To lngDim)
            lngDesign(lngDim) = 2
            strDesign(lngDim) = FormInputs.chkDesignTight.Caption
        End If

        
        '---------Download CSV with design specs from Confluence site-------
        Dim arrDesign() As Variant
        Dim strCastoffFile As String    'File name of CSV on Confluence
    
        'Create name of castoff csv file to download
        strCastoffFile = "Castoff_" & strPub & ".csv"
        
        arrDesign = DownloadCSV(FileName:=strCastoffFile, Staging:=blnStaging)
        
        ' Check that returned array is allocated
        If IsArrayEmpty(arrDesign) = True Then
            Exit Sub ' Error messages were in DownloadCSV (and DownloadFromConfluence) so none needed here
        End If
        
        '------------Get castoff for each Design selected-------------------
        Dim lngTrimIndex As Long
        If FormInputs.optTrim5x8 Then  ' already validated that there is at least 1 picked in form code
            lngTrimIndex = 0
        Else
            lngTrimIndex = 1
        End If
        
        Dim d As Long
        
        For d = LBound(lngDesign()) To UBound(lngDesign())
            'Debug.Print _
            UBound(arrDesign(), 1) & " < " & lngDesign(d) & vbNewLine & _
            UBound(arrDesign(), 2) & " < "; intTrim
            
            'Error handling: lngDesign(d) must be in range of design array
            If UBound(arrDesign(), 1) <= lngDesign(d) And UBound(arrDesign(), 2) <= lngTrimIndex Then
                 MsgBox "There was an error generating your castoff. Please contact workflows@macmillan.com for assistance.", _
                    vbCritical, "Error 1: Design Count Out of Range"
                Unload FormInputs
                Exit Sub
            Else
    
                '---------Calculate Page Count--------------------------------------
                ReDim Preserve lngCastoffResult(d)
                lngCastoffResult(d) = Castoff(lngDesign(d), arrDesign(), FormInputs, blnStaging)
                
            End If
        Next d
        
        ' ----- Get spine size if POD -------
        Dim strSpineSize As String
        strSpineSize = ""
        
        If FormInputs.optPrintPOD Then
            strSpineSize = SpineSize(blnStaging, lngCastoffResult(0))
            'Debug.Print "spine size = " & strSpineSize
        End If
    End If
    
    '-------------Create final message---------------------------------------------------
    Dim strReportText As String
    
    'Get Title Information from Form
    Dim strEditor As String
    strEditor = FormInputs.txtEditor.value
    
    Dim strAuthor As String
    strAuthor = FormInputs.txtAuthor.value
    
    Dim strTitle As String
    strTitle = FormInputs.txtTitle.value
    
    Dim lngSchedPgCount As Long
    lngSchedPgCount = FormInputs.numTxtPageCount.value
    
    ' Get trim size (here and not above because need to do for Pickup)
    Dim strTrimSize As String
    If FormInputs.optTrim5x8 Then
        strTrimSize = FormInputs.optTrim5x8.Caption
    ElseIf FormInputs.optTrim6x9 Then
        strTrimSize = FormInputs.optTrim6x9.Caption
    End If
    
    ' Get print type from Form
    Dim strPrintType As String
    If FormInputs.optPrintOffset Then
        strPrintType = FormInputs.optPrintOffset.Caption
    Else
        strPrintType = FormInputs.optPrintPOD.Caption
    End If
    
    ' Create text of castoff from arrays
    Dim strCastoffs As String
    Dim strPickupTitle As String
    Dim e As Long
    
    ' If it's a pickup, there is only 1 option
    If FormInputs.optPubPickup Then
        strCastoffs = lngCastoffResult(0)
        strPickupTitle = "PICKUP TITLE: " & FormInputs.txtPrevTitle.value & vbNewLine
    Else
        strPickupTitle = ""
        strCastoffs = vbNewLine
        For e = LBound(lngCastoffResult) To UBound(lngCastoffResult)
            strCastoffs = strCastoffs & vbTab & strDesign(e) & ": " & lngCastoffResult(e) & vbNewLine
        Next e
    End If
    
    
    strReportText = _
    " * * * MACMILLAN PRELIMINARY CASTOFF * * * " & vbNewLine & _
    vbNewLine & _
    "DATE: " & Date & vbNewLine & _
    "TITLE: " & strTitle & vbNewLine & _
    "AUTHOR: " & strAuthor & vbNewLine & _
    "PUBLISHER: " & strPubRealName & vbNewLine & _
    "EDITOR: " & strEditor & vbNewLine & _
    "PRINTING TYPE: " & strPrintType & vbNewLine & _
    "TRIM SIZE: " & strTrimSize & vbNewLine & _
    vbNewLine & _
    strPickupTitle & _
    "SCHEDULED PAGE COUNT: " & lngSchedPgCount & vbNewLine & _
    "ESTIMATED PAGE COUNT: " & _
    strCastoffs & _
    vbNewLine & _
    strSpineSize
    
    '-------------Report castoff info to user----------------------------------------------------------------
    Call CreateTextFile(strText:=strReportText, suffix:="Castoff")

    Unload FormInputs
            
End Sub

Private Function LoadCSVtoArray(Path As String, RemoveHeaderRow As Boolean, RemoveHeaderCol As Boolean) As Variant

'------Load CSV into 2d array, NOTE!!: base 0---------
' But also note that this now removes the header row and column too
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
        
        ' Do we need to remove a header row?
        Dim lngHeaderRow As Long
        If RemoveHeaderRow = True Then
            lngHeaderRow = 1
        Else
            lngHeaderRow = 0
        End If
        
        ' Do we need to remove a header column?
        Dim lngHeaderCol As Long
        If RemoveHeaderCol = True Then
            lngHeaderCol = 1
        Else
            lngHeaderCol = 0
        End If
        
        ' Load the csv file.
        fnum = FreeFile
        Open Path For Input As fnum
        whole_file = Input$(LOF(fnum), #fnum)
        Close fnum

        ' Break the file into lines (trying to capture whichever line break is used)
        If InStr(1, whole_file, vbCrLf) <> 0 Then
            lines = Split(whole_file, vbCrLf)
        ElseIf InStr(1, whole_file, vbCr) <> 0 Then
            lines = Split(whole_file, vbCr)
        ElseIf InStr(1, whole_file, vbLf) <> 0 Then
            lines = Split(whole_file, vbLf)
        Else
            MsgBox "There was an error with your castoff.", vbCritical, "Error parsing CSV file"
        End If

        ' Dimension the array.
        num_rows = UBound(lines)
        one_line = Split(lines(0), ",")
        num_cols = UBound(one_line)
        ReDim the_array(num_rows - lngHeaderRow, num_cols - lngHeaderCol) ' -1 if we are not using header row/col
        
        ' Copy the data into the array.
        For R = lngHeaderRow To num_rows           ' start at 1 (not 0) if we are not using the header row
            If Len(lines(R)) > 0 Then
                one_line = Split(lines(R), ",")
                For c = lngHeaderCol To num_cols   ' start at 1 (not 0) if we are not using the header column
                    the_array((R - lngHeaderRow), (c - lngHeaderCol)) = one_line(c)   ' -1 because if are not using header row/column from CSV
                Next c
            End If
        Next R
    
        ' Prove we have the data loaded.
        ' Debug.Print LBound(the_array)
        ' Debug.Print UBound(the_array)
        ' For R = 0 To num_rows          ' -1 again if we removed the header row
        '     For c = 0 To num_cols      ' -1 again if we removed the header column
        '         Debug.Print the_array(R, c) & "|";
        '     Next c
        '     Debug.Print
        ' Next R
        ' Debug.Print "======="
    
    LoadCSVtoArray = the_array
    
End Function

Private Function Castoff(lngDesignIndex As Long, arrCSV() As Variant, objForm As CastoffForm, StagingThing As Boolean) As Long
    
    ' Get total character count (incl. notes and spaces) from document
    Dim lngTotalCharCount As Long
    lngTotalCharCount = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticCharactersWithSpaces, _
                        IncludeFootnotesAndEndnotes:=True)
                        
    ' Get char count for just embedded endnotes and footnotes
    Dim lngNotesCharCount As Long
    lngNotesCharCount = lngTotalCharCount - ActiveDocument.ComputeStatistics(Statistic:=wdStatisticCharactersWithSpaces, _
                        IncludeFootnotesAndEndnotes:=False)
          
    ' Get page count of main text
    Dim lngMainTextPages As Long
    ActiveDocument.Repaginate
    lngMainTextPages = ActiveDocument.StoryRanges(wdMainTextStory).Information(wdNumberOfPagesInDocument)
        
    ' Get page count of endnotes and footnotes
    Dim lngEndnotesPages As Long
    Dim lngFootnotesPages As Long
    Dim lngNotesPages As Long
    
    If EndnotesExist = True Then
        ActiveDocument.Repaginate
        lngEndnotesPages = ActiveDocument.StoryRanges(wdEndnotesStory).Information(wdNumberOfPagesInDocument)
    Else
        lngEndnotesPages = 0
    End If
    
    If FootnotesExist = True Then
        ActiveDocument.Repaginate
        lngFootnotesPages = ActiveDocument.StoryRanges(wdFootnotesStory).Information(wdNumberOfPagesInDocument)
    Else
        lngFootnotesPages = 0
    End If
    
    lngNotesPages = lngEndnotesPages + lngFootnotesPages
                    
    ' Calculate number of characters per page of NOTES in the MANUSCRIPT if there are linked notes
    ' If there aren't linked notes, get char/page of whole manuscript
    Dim lngMsCharPerPage As Long
    If lngNotesPages > 0 Then
        lngMsCharPerPage = lngNotesCharCount / lngNotesPages
    Else
        lngMsCharPerPage = lngTotalCharCount / lngMainTextPages
    End If
    
    ' Get number of unlinked endnotes in MS from Form, estimate number of characters
    Dim lngUnlinkedNotesCharCount As Long
    If objForm.numTxtUnlinkedNotes <> vbNullString Then
        lngUnlinkedNotesCharCount = objForm.numTxtUnlinkedNotes * lngMsCharPerPage
    Else
        lngUnlinkedNotesCharCount = 0
    End If
    
    ' Get number of endnotes TK from Form, estimate number of characters
    Dim lngEndnotesTKCharCount As Long
    If objForm.numTxtNotesTK <> vbNullString Then
        lngEndnotesTKCharCount = objForm.numTxtNotesTK * lngMsCharPerPage
    Else
        lngEndnotesTKCharCount = 0
    End If
    
    ' Get number of biblio pages in manuscript from Form, estimate number of characters
    Dim lngBiblioMsCharCount As Long
    If objForm.numTxtBibliography <> vbNullString Then
        lngBiblioMsCharCount = objForm.numTxtBibliography * lngMsCharPerPage
    Else
        lngBiblioMsCharCount = 0
    End If
    
    ' Get number of biblio pages TK from Form, estimate number of characters
    Dim lngBiblioTKCharCount As Long
    If objForm.numTxtBiblioTK <> vbNullString Then
        lngBiblioTKCharCount = objForm.numTxtBiblioTK * lngMsCharPerPage
    Else
        lngBiblioTKCharCount = 0
    End If

    
    ' Calculate total character count of main text and notes separately
    Dim lngMainCharCount As Long
    lngMainCharCount = lngTotalCharCount - lngNotesCharCount - lngUnlinkedNotesCharCount - lngBiblioMsCharCount
    lngNotesCharCount = lngNotesCharCount + lngUnlinkedNotesCharCount + lngBiblioMsCharCount + lngEndnotesTKCharCount _
                        + lngBiblioTKCharCount
                        
    ' Get trim size from Form
    ' Number assigned is column index in design array
    ' 0 = 5-1/2 x 8-1/4
    ' 1 = 6-1/8 x 9-1/4
    
    Dim lngTrim As Long
    If objForm.optTrim5x8 Then  ' already validated that there is at least 1 picked in form code
        lngTrim = 0
    Else
        lngTrim = 1
    End If
    
    ' --------------------------------------------------
    ' For Reference: Index numbers in arrCSV (base 0)
    '
    '         | 5-1/2 x 8-1/4 |  6-1/8 x 9-1/4
    'loose    | (0,0)         | (0,1)
    'average  | (1,0)         | (1,1)
    'tight    | (2,0)         | (2,1)
    'notes    | (3,0)         | (3,1)
    'lines    | (4,0)         | (4,1)
    'overflow | (5,0)         | (5,1)
    '--------------------------------------------------
    
    '---------Get design character count from CSV-------------------------------
    Dim lngDesignCount As Long
    lngDesignCount = arrCSV(lngDesignIndex, lngTrim)
    'Debug.Print lngDesignCount
    
    '---------Get notes character count from CSV--------------------------------
    Dim lngNotesDesign As Long
    lngNotesDesign = arrCSV(3, lngTrim)
    
    '---------Get lines per page from CSV--------------------------------------
    Dim lngLinesPage As Long
    lngLinesPage = arrCSV(4, lngTrim)
    
    '---------Get overflow pages from CSV--------------------------------------
    Dim lngOverflow As Long
    lngOverflow = arrCSV(5, lngTrim)

    '----------Get user inputs from Userform--------------------------------------------------
    ' Get info from Standard Items section (already validated as having data)
    Dim lngChapters As Long      ' number of chapters
    lngChapters = objForm.numTxtChapters
    
    Dim lngParts As Long         'number of part titles
    lngParts = objForm.numTxtParts
    
    Dim lngFMpgs As Long         ' number of pages of frontmatter including blanks
    lngFMpgs = objForm.numTxtFrontmatter
    
    ' The rest of the inputs are not required, so only assign the value if one exists
    ' Otherwise assign 0, so we can still use the variable later without a whole other
    ' bunch of if statements
    
    ' Get info from Back Matter section
    Dim lngIndexPgs As Long     'Number of pages estimated for the index
    If objForm.numTxtIndex <> vbNullString Then
        lngIndexPgs = objForm.numTxtIndex
    Else
        lngIndexPgs = 0
    End If
    
    Dim lngBackmatterPgsTK As Long 'Number of pages of backmatter TK
    If objForm.numTxtBackmatter <> vbNullString Then
        lngBackmatterPgsTK = objForm.numTxtBackmatter
    Else
        lngBackmatterPgsTK = 0
    End If
    
    ' Get info from Complex Items section
    Dim lngSubheads2Chap As Long 'Number of subheads in 2 chapters
    If objForm.numTxtSubheads <> vbNullString Then
        lngSubheads2Chap = objForm.numTxtSubheads
    Else
        lngSubheads2Chap = 0
    End If
    
    Dim lngTablesPgs As Long  'Number of pages for tables
    If objForm.numTxtTables <> vbNullString Then
        lngTablesPgs = objForm.numTxtTables
    Else
        lngTablesPgs = 0
    End If
    
    Dim lngArtPgs As Long  'Number of pages for in-text art
    If objForm.numTxtArt <> vbNullString Then
        lngArtPgs = objForm.numTxtArt
    Else
        lngArtPgs = 0
    End If

    ' Calculate number of pages!
    Dim lngMainPages As Long
    Dim lngTotalNotesPages As Long
    Dim lngPartsPages As Long
    Dim lngHeadingPages As Long
    Dim lngEstPages As Long
    
    lngMainPages = lngMainCharCount / lngDesignCount
    lngTotalNotesPages = lngNotesCharCount / lngNotesDesign
    lngPartsPages = lngParts * 2
    lngHeadingPages = ((lngSubheads2Chap / 2) * lngChapters * 3) / lngLinesPage  ' 3 because headings take up 3 lines each
    
    lngEstPages = lngMainPages _
                + lngTotalNotesPages _
                + lngPartsPages _
                + lngHeadingPages _
                + lngChapters _
                + lngFMpgs _
                + lngIndexPgs _
                + lngBackmatterPgsTK _
                + lngTablesPgs _
                + lngArtPgs
                
    Dim lngFinalResult As Long
    
    lngFinalResult = FinalSig(lngEstPages, objForm, StagingThing)
    
    Castoff = lngFinalResult

End Function
Private Function SpineSize(StagingSite As Boolean, PageCount As Long)
    Dim strSpine As String
    
    If PageCount < 48 Then
        strSpine = "NOTE: POD titles less than 48 pages will be saddle-stitched."
    ElseIf PageCount >= 48 And PageCount <= 1050 Then       'Limits of spine size table
    
        '----Define spine chart file name--------------------------------
        Dim strSpineFile As String
        'strSpineFile = "Spine_" & Publisher & ".csv"   ' Use this if we're doing different paper based on imprint
        strSpineFile = "POD_Spines.csv"                 ' Use this if we're doing 1 kind of paper for all POD
        
        Dim arrSpine() As Variant
        
        arrSpine = DownloadCSV(FileName:=strSpineFile, Staging:=StagingSite)
        
        ' Check that returned array is allocated
        If IsArrayEmpty(arrSpine) = True Then
            Exit Function ' Error messages were in DownloadCSV (and DownloadFromConfluence) so none needed here
        End If
    
        '---------Lookup spine size in array-------------------------------
        Dim c As Long
        
        For c = LBound(arrSpine, 1) To UBound(arrSpine, 1)
            'Debug.Print arrDesign(c, 0) & " = " & PageCount
            If arrSpine(c, 0) = PageCount Then
                strSpine = arrSpine(c, 1)
                Exit For
            End If
        Next c
    Else
        strSpine = "Your page count of " & PageCount & _
                " is out of range of the spine-size table."
    End If
    
    
    If strSpine = vbNullString Then
        strSpine = "Word was unable to generate a spine size. " & _
            "Contact workflows@macmillan.com for assistance."
    Else
        strSpine = "Your spine size will be " & strSpine & " inches " & _
                            "at this page count."
    End If

    'Debug.Print strSpine
    SpineSize = strSpine

End Function

Private Function PickupDesign(objCastoffForm As CastoffForm) As Long
' estimate page count based on design of previous book
    
    ' Get total character count of pickup book from form
    Dim lngPrevMsCharCount As Long
    lngPrevMsCharCount = objCastoffForm.numTxtPrevCharCount.value
    
    ' get final page count of pickup book from form
    Dim lngPrevBookPageCount As Long
    lngPrevBookPageCount = objCastoffForm.numTxtPrevPageCount.value
    
    ' get additional pages from form
    Dim lngAddlPages As Long
    lngAddlPages = objCastoffForm.numTxtAddlPgs.value
    
    ' get total character count of current ms from document
    Dim lngCurrentMsCharCount As Long
    lngCurrentMsCharCount = ActiveDocument.ComputeStatistics(Statistic:=wdStatisticCharactersWithSpaces, _
                        IncludeFootnotesAndEndnotes:=True)
    
    ' divide total prev character count by page count to get avg characters per page in prev book
    Dim lngPrevCharPerBookPage As Long
    lngPrevCharPerBookPage = lngPrevMsCharCount / lngPrevBookPageCount
    
    ' divide total characters of this doc by avg characters per page to get est page count, add additional pages
    Dim lngStartingResult As Long
    lngStartingResult = (lngCurrentMsCharCount / lngPrevCharPerBookPage) + lngAddlPages
    
    ' Calculate what the final sig will be
    Dim lngFinalPageCount As Long
    lngFinalPageCount = FinalSig(lngStartingResult, objCastoffForm)
    
    PickupDesign = lngFinalPageCount
    
End Function

    
Private Function FinalSig(RawEstPages As Long, objCastForm As CastoffForm, StagingUsed As Boolean) As Long
    ' Figure out what the final sig/page count will be
    Dim result As Long
           
    If objCastForm.optPrintPOD Then
        'POD only has to be even, not 16-page sig
        If (RawEstPages Mod 2) = 0 Then      'page count is even
            result = RawEstPages
        Else                                    'page count is odd
            result = RawEstPages + 1
        End If
    Else 'It's printing offset, already validated in castoff form code
        
        ' Calculate next sig up and next sig down
        Dim lngRemainderPgs As Long
        Dim lngLowerSig As Long
        Dim lngUpperSig As Long
        
        lngRemainderPgs = RawEstPages Mod 16
        lngLowerSig = RawEstPages - lngRemainderPgs
        lngUpperSig = RawEstPages + (16 - lngRemainderPgs)
                    
        ' Get number of overflow pages from CSV
        Dim arrCastoff() As Variant
        Dim strFile As String
        
        If objCastForm.optPubSMP Or objCastForm.optPubPickup Then
            strFile = "Castoff_SMP.csv"
        Else
            strFile = "Castoff_TorDOTcom.csv"
        End If
        
        Dim lngTrimIndex As Long
        If objCastForm.optTrim5x8 Then
            lngTrimIndex = 0
        Else
            lngTrimIndex = 1
        End If
        
        arrCastoff = DownloadCSV(FileName:=strFile, Staging:=StagingUsed)
        
        Dim lngOverflow As Long
        lngOverflow = arrCastoff(5, lngTrimIndex)    ' 5 is index of overflow info in CSV
        
        ' Determine if we go up or down a signature
        If lngRemainderPgs < lngOverflow Then    ' Do we want this value in a CSV on Confluence for easy update?
            result = lngLowerSig
        Else
            result = lngUpperSig
        End If
    End If

    FinalSig = result
    
End Function


Private Function DownloadCSV(FileName As String, Staging As Boolean) As Variant
    '---------Download CSV with design specs from Confluence site-------

    'Create log file name
    Dim arrLogInfo() As Variant
    ReDim arrLogInfo(1 To 3)
    
    arrLogInfo() = CreateLogFileInfo(FileName)
      
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
    strPath = strDir & Application.PathSeparator & FileName
        
    'Check if log file already exists; if not, create it
    CheckLog strStyleDir, strDir, strLogFile
    
    'Download CSV file from Confluence
    If DownloadFromConfluence(Staging, strDir, strLogFile, FileName) = False Then
        ' If download fails, check if we have an older version of the CSV to work with
        If IsItThere(strPath) = False Then
            strMessage = "Looks like we can't download the design info from the internet right now. " & _
                "Please check your internet connection, or contact workflows@macmillan.com."
            MsgBox strMessage, vbCritical, "Error 5: Download failed, no previous design file available"
            Exit Function
        Else
            strMessage = "Looks like we can't download the most up to date design info from the internet right now, " & _
                "so we'll just use the info we have on file for your castoff."
            MsgBox strMessage, vbInformation, "Let's do this thing!"
        End If
    End If
    
    ' Heading row/col different based on different InfoTypes
    Dim blnRemoveHeaderRow As Boolean
    Dim blnRemoveHeaderCol As Boolean
    
    If InStr(1, FileName, "Castoff") <> 0 Then
        blnRemoveHeaderRow = True
        blnRemoveHeaderCol = True
    ElseIf InStr(1, FileName, "Spine") <> 0 Then
        blnRemoveHeaderRow = True
        blnRemoveHeaderCol = False
    End If
    
    'Double check that CSV is there
    Dim arrFinal() As Variant
    If IsItThere(strPath) = False Then
        strMessage = "The Castoff macro is unable to access the design count file right now. Please check your internet " & _
                    "connection and try again, or contact workflows@macmillan.com."
        MsgBox strMessage, vbCritical, "Error 3: Design CSV doesn't exist"
        Exit Function
    Else
        ' Load CSV into an array
        arrFinal = LoadCSVtoArray(Path:=strPath, RemoveHeaderRow:=blnRemoveHeaderRow, RemoveHeaderCol:=blnRemoveHeaderCol)
    End If
    
    DownloadCSV = arrFinal
    
End Function
