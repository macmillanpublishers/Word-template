Attribute VB_Name = "CastoffMacro"
Option Explicit
Sub UniversalCastoff()
'created by Erica Warren - erica.warren@macmillan.com

    '----------Load userform to get user inputs------------------------
    Dim objCastoffForm As CastoffForm
    Set objCastoffForm = New CastoffForm
    
    objCastoffForm.Show
    
    If objCastoffForm.blnCancel = True Then
        Unload objCastoffForm
        Exit Sub
    End If

    '----------Get user inputs from Userform---------------------------
    Dim intTrim As Integer
    Dim strTrim As String
    Dim intDesign() As Integer
    Dim strDesign() As String
    Dim intDim As Integer
    Dim strPub As String

    'Debug.Print objCastoffForm.tabPublisher.SelectedItem.Caption

    'Get trim size.
    'Number assigned is column index in design array
    '0 = 5-1/2 x 8-1/4
    '1 = 6-1/8 x 9-1/4
    If objCastoffForm.optTrim5x8 Then
        intTrim = 0
        strTrim = "5-1/2 x 8-1/4"
    ElseIf objCastoffForm.optTrim6x9 Then
        intTrim = 1
        strTrim = "6-1/8 x 9-1/4"
    Else
        MsgBox "You must select a Trim Size to run the Castoff Macro."
        objCastoffForm.Show
    End If
        
    'Get designs selected.
    'Number for intDesign is row index in design array
    intDim = 0
    
    If objCastoffForm.chkDesignLoose Then
        intDim = intDim + 1
        ReDim Preserve intDesign(1 To intDim)
        ReDim Preserve strDesign(1 To intDim)
        intDesign(intDim) = 0
        strDesign(intDim) = objCastoffForm.chkDesignLoose.Caption
    End If
    
    If objCastoffForm.chkDesignAverage Then
        intDim = intDim + 1
        ReDim Preserve intDesign(1 To intDim)
        ReDim Preserve strDesign(1 To intDim)
        intDesign(intDim) = 1
        strDesign(intDim) = objCastoffForm.chkDesignAverage.Caption
    End If
    
    If objCastoffForm.chkDesignTight Then
        intDim = intDim + 1
        ReDim Preserve intDesign(1 To intDim)
        ReDim Preserve strDesign(1 To intDim)
        intDesign(intDim) = 2
        strDesign(intDim) = objCastoffForm.chkDesignTight.Caption
    End If
    
    'Make sure at least one design is selected
    If intDim = 0 Then
        MsgBox "You must select at least one Design to run the Castoff Macro."
        objCastoffForm.Show
    End If
    
    'Get publisher name from tab of userform
    strPub = objCastoffForm.tabPublisher.SelectedItem.Caption
    
    '---------Download CSV with design specs from Confluence site-------

    'Need separate PC and Mac subs to download file
    Dim TheOS As String
    Dim strPath As String
    Dim strInfoType As String

    TheOS = System.OperatingSystem
    strInfoType = "Castoff"

    If Not TheOS Like "*Mac*" Then
        strPath = GetCSV_PC(strInfoType, strPub)
            If strPath = vbNullString Then
                MsgBox "The Castoff Macro can't access the source design count file right now. Please check your internet connection.", _
                    vbCritical, "Error 3: Path to CSV Is Null"
                Unload objCastoffForm
                Exit Sub
            End If
    Else
        strPath = GetCSV_Mac(strInfoType, strPub)
            If strPath = vbNullString Then
                MsgBox "The Castoff Macro can't access the source design count file right now. Please check your internet connection.", _
                    vbCritical, "Error 3: Path to CSV Is Null"
                Unload objCastoffForm
                Exit Sub
            End If
    End If

    '---------Load CSV into an array-----------------------------------
    Dim arrDesign() As Variant
    arrDesign = LoadCSVtoArray(strPath)

            
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
            Unload objCastoffForm
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

            arrCastoffResult = Castoff(lngDesignCount, objCastoffForm)

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
            
                Debug.Print strPub
                
                'Get spine size
                If lngFinalCount >= 18 And lngFinalCount <= 1050 Then       'Limits of spine size table
                    strSpineSize = SpineSize(lngFinalCount, strPub, objCastoffForm)
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
            Dim strMessage As String
            
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

    Unload objCastoffForm
            
End Sub

Private Function GetCSV_PC(InfoType As String, Publisher As String) As String

    Dim WinHttpReq As Object
    Dim oStream As Object
    Dim strCastoffURL As String
    Dim strCastoffFile As String
    Dim dirNamePC As String
    Dim myURL As String
    
    'this is download link, actual page housing file is https://confluence.macmillan.com/display/PBL/Test
    strCastoffURL = "https://confluence.macmillan.com/download/attachments/9044274/"
    'CSV on Confluence page must match this format:
    strCastoffFile = InfoType & "_" & Publisher & ".csv"
    myURL = strCastoffURL & strCastoffFile
    dirNamePC = Environ("TEMP") & "\" & strCastoffFile

    'Debug.Print dirNamePC
    
    'Attempt to download file
    On Error Resume Next
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False
    WinHttpReq.Send
    While WinHttpReq.readyState <> 4
        DoEvents
    Wend

        ' Exit if error in connecting to website
        If Err.Number <> 0 Then 'HTTP request is not OK
            GetCSV_PC = ""
            Exit Function
        End If
    On Error GoTo 0

    'Debug.Print WinHttpReq.Status

    If WinHttpReq.Status = 200 Then  ' 200 = HTTP request is OK
    
        'if connection OK, download file and save to log directory
        strCastoffURL = WinHttpReq.responseBody
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile dirNamePC, 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
        Set oStream = Nothing
        Set WinHttpReq = Nothing
    Else
        GetCSV_PC = ""
        Exit Function
    End If
    
    'Check if download was successful
    If Dir(dirNamePC) = vbNullString Then
        GetCSV_PC = ""
        Exit Function
    End If
    
    GetCSV_PC = dirNamePC

End Function

Private Function GetCSV_Mac(InfoType As String, Publisher As String) As String
    Dim dirNameMac As String
    Dim dirNameBash As String
    Dim strCastoffFile As String
    Dim dlUrl As String
    
    dirNameMac = "Macintosh HD:private:tmp:"
    dirNameBash = "/private/tmp/"
    strCastoffFile = InfoType & "_" & Publisher & ".csv"
    dlUrl = "https://confluence.macmillan.com/download/attachments/9044274/"
    
    'Debug.Print strCastoffFile
    
    'check for network.  Skipping domain since we are looking at confluence, but would test ping hbpub.net or mpl.root-domain.org
    If ShellAndWaitMac("ping -o google.com &> /dev/null ; echo $?") <> 0 Then
        GetCSV_Mac = ""
        Exit Function
    End If
    
    'download CSV file to temp
    ShellAndWaitMac ("rm -f " & dirNameBash & strCastoffFile & " ; curl -o " & dirNameBash & strCastoffFile & " " & dlUrl & strCastoffFile)
    
    'return full path to CSV file
    GetCSV_Mac = dirNameMac & strCastoffFile
    
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

        'For R = 0 To num_rows
        '    For c = 0 To num_cols
        '        Debug.Print the_array(R, c) & "|";
        '    Next c
        '    Debug.Print
        'Next R
        'Debug.Print "======="
        
        'Delete the .csv file 'cuz we don't need it any more
        If Len(Dir$(Path)) > 0 Then
            Kill Path
        End If
    
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
Private Function SpineSize(PageCount As Long, Publisher As String, objForm As CastoffForm)

'----Download CSV with spine sizes from Confluence site----------

    'Need separate PC and Mac subs to download file
    Dim TheOS As String
    Dim strPath As String
    Dim strInfoType As String
    Dim strPub As String

    TheOS = System.OperatingSystem
    strInfoType = "Spine"

    If Not TheOS Like "*Mac*" Then
        strPath = GetCSV_PC(strInfoType, Publisher)
            If strPath = vbNullString Then
                MsgBox "The Castoff Macro can't access the source spine size file right now. Please check your internet connection."
                Unload objForm
                Exit Function
            End If
    Else
        strPath = GetCSV_Mac(strInfoType, Publisher)
            If strPath = vbNullString Then
                MsgBox "The Castoff Macro can't access the source spine size file right now. Please check your internet connection."
                Unload objForm
                Exit Function
            End If
    End If

'---------Load CSV into an array-----------------------------------
    Dim arrDesign() As Variant
    arrDesign = LoadCSVtoArray(strPath)
    
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
Private Function ShellAndWaitMac(cmd As String) As String

Dim result As String
Dim scriptCmd As String ' Macscript command
'
scriptCmd = "do shell script """ & cmd & """"
result = MacScript(scriptCmd) ' result contains stdout, should you care
ShellAndWaitMac = result
End Function
