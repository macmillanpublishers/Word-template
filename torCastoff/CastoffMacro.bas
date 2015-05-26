Attribute VB_Name = "CastoffMacro"
Option Explicit
Sub UniversalCastoff()

'----------Load userform to get user inputs------------------------
    Load CastoffForm
    CastoffForm.Show
    
    If CastoffForm.blnCancel = True Then
        Unload CastoffForm
        Exit Sub
    End If

'----------Get user inputs from Userform---------------------------
    Dim intTrim As Integer
    Dim intDesign As Integer
    Dim strPub As String

    Debug.Print CastoffForm.tabPublisher.SelectedItem.Caption

    'Get trim size.
    '0 = 5-1/2 x 8-1/4
    '1 = 6-1/8 x 9-1/4
    If CastoffForm.optTrim5x8 Then
        intTrim = 0
    ElseIf CastoffForm.optTrim6x9 Then
        intTrim = 1
    Else
        MsgBox "You must select a Trim Size to run the Castoff Macro."
        CastoffForm.Show
    End If
        
    'Get design.
    '0 = loose
    '1 = average
    '2 = tight
    If CastoffForm.optDesignLoose Then
        intDesign = 0
    ElseIf CastoffForm.optDesignAverage Then
        intDesign = 1
    ElseIf CastoffForm.optDesignTight Then
        intDesign = 2
    Else
        MsgBox "You must select a Design to run the Castoff Macro."
        CastoffForm.Show
    End If
    
    'Get publisher name from tab of userform
    strPub = CastoffForm.tabPublisher.SelectedItem.Caption
    
'---------Download CSV with design specs from Confluence site-------

    'Need separate PC and Mac subs to download file
    Dim TheOS As String
    Dim strPath As String

    TheOS = System.OperatingSystem

    If Not TheOS Like "*Mac*" Then
        strPath = GetCSV_PC(strPub)
            If strPath = vbNullString Then
                MsgBox "The Castoff Macro can't access the source design count file right now. Please check your internet connection."
                Exit Sub
            End If
    Else
        strPath = GetCSV_Mac(strPub)
            If strPath = vbNullString Then
                MsgBox "The Castoff Macro can't access the source design count file right now. Please check your internet connection."
                Exit Sub
            End If
    End If

'---------Load CSV into an array-----------------------------------
    Dim arrDesign() As Variant
    arrDesign = LoadCSVtoArray(strPath)

            
'---------Get design character count-------------------------------
    Dim lngDesignCount As Long

    lngDesignCount = arrDesign(intDesign, intTrim)
    
    '--------------------------------------------------
    'For Reference: Index numbers in array (base 0)
    '
    '       | 5-1/2 x 8-1/4 |  6-1/8 x 9-1/4
    'loose  | (0,0)         | (0,1)
    'average| (1,0)         | (1,1)
    'tight  | (2,0)         | (2,1)
    '--------------------------------------------------

    Debug.Print lngDesignCount

'---------Calculate Page Count--------------------------------------
    Dim arrCastoffResult() As Variant
    Dim lngFinalCount As Long
    Dim lngBlankPgs As Long
    Dim lngActualCount As Long

    arrCastoffResult = Castoff(lngDesignCount)

    lngFinalCount = arrCastoffResult(0)
    lngBlankPgs = arrCastoffResult(1)
    lngActualCount = arrCastoffResult(2)

    Debug.Print lngFinalCount

    MsgBox "Your book will be approximately " & lngFinalCount & " pages." & vbCr & vbCr & _
            vbTab & lngActualCount & " text pages" & vbCr & _
            vbTab & "  " & lngBlankPgs & " blank pages" & vbCr & _
            vbTab & "____________________" & vbCr & _
            vbTab & lngFinalCount & " total pages"
            
End Sub

Private Function GetCSV_PC(Publisher As String) As String

    Dim WinHttpReq As Object
    Dim oStream As Object
    Dim strCastoffURL As String
    Dim strCastoffFile As String
    Dim dirNamePC As String
    Dim myURL As String
    
    'this is download link, actual page housing file is http://confluence.macmillan.com/display/PBL/Test
    strCastoffURL = "http://confluence.macmillan.com/download/attachments/9044274/"
    strCastoffFile = "Castoff_" & Publisher & ".csv"
    myURL = strCastoffURL & strCastoffFile
    dirNamePC = Environ("TEMP")

    Debug.Print dirNamePC
    
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False
    On Error Resume Next
    WinHttpReq.Send

        ' Exit sub if error in connecting to website
        If Err.Number <> 0 Then 'HTTP request is not OK
            GetCSV_PC = ""
            Exit Function
        End If
    On Error GoTo 0

    Debug.Print WinHttpReq.Status

    If WinHttpReq.Status = 200 Then  ' 200 = HTTP request is OK
    
        'if connection OK, download file and save to log directory
        strCastoffURL = WinHttpReq.responseBody
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile dirNamePC & "\" & strCastoffFile, 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
        Set oStream = Nothing
        Set WinHttpReq = Nothing
    End If

    Dim strFinalPath As String

    'Change this to where .xls will be downloaded to
    strFinalPath = dirNamePC & "\" & strCastoffFile

    GetCSV_PC = strFinalPath

End Function

Private Function GetCSV_Mac()

End Function

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

        For R = 0 To num_rows
            For c = 0 To num_cols
                Debug.Print the_array(R, c) & "|";
            Next c
            Debug.Print
        Next R
        Debug.Print "======="
        
        'Delete the .csv file 'cuz we don't need it any more
        If Len(Dir$(Path)) > 0 Then
            Kill Path
        End If
    
    LoadCSVtoArray = the_array
    
End Function

Private Function Castoff(Design As Long) As Variant
    Dim lngCharacterCount As Long
    Dim lngActualPageCount As Long
    Dim lngFinalPageCount As Long

    'Get character count with space from Word doc, divide by avg. char. count of design to get page count
    lngCharacterCount = ActiveDocument.Range.ComputeStatistics(wdStatisticCharactersWithSpaces)
    lngActualPageCount = lngCharacterCount / Design

    Debug.Print lngActualPageCount
    
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

    Debug.Print lngFinalPageCount

    Dim lngBlanks As Long
    Dim lngPgsOver As Long
    lngBlanks = 0
    lngPgsOver = (lngFinalPageCount Mod 16)

    If lngPgsOver <> 0 Then
        lngFinalPageCount = lngFinalPageCount + (16 - lngPgsOver)
        lngBlanks = 16 - lngPgsOver
    Else
        lngBlanks = 0
    End If

    Dim arrResult() As Variant
    ReDim arrResult(0 To 2)

    arrResult(0) = lngFinalPageCount
    arrResult(1) = lngBlanks
    arrResult(2) = lngFinalPageCount - lngBlanks

    Castoff = arrResult

End Function
