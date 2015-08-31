Attribute VB_Name = "Endnotes"
Option Explicit
Option Base 1
Dim activeRng As Range

Sub EndnoteDeEmbed()
Dim refRng As Range
Dim refSection As Integer
Dim lastRefSection As Integer
Dim chapterName As String
Dim addChapterName As Boolean
Dim addHeader As Boolean
Dim nRng As Range, eNote As Endnote, nref As String, refCopy As String
Dim sectionCount As Long
Dim StoryRange As Range
Dim EndnotesExist As Boolean
Dim TheOS As String
Dim palgraveTag As Boolean
Dim iReply As Integer
Dim BookmarkNum As Integer
Dim BookmarkName As String

BookmarkNum = 1
lastRefSection = 0
addHeader = True
EndnotesExist = False
TheOS = System.OperatingSystem
palgraveTag = False

'''Error checks, setup Doc with sections & numbering
If TheOS Like "*Mac*" Then
    MsgBox "Mac OS detected.  This macro will not work properly in Word for Mac. Exiting Endnotes macro."
    Exit Sub
End If
For Each StoryRange In ActiveDocument.StoryRanges
    If StoryRange.StoryType = wdEndnotesStory Then
        EndnotesExist = True
        Exit For
    End If
Next StoryRange
If EndnotesExist = False Then
    MsgBox "No endnotes found in document.  Exiting Endnotes macro."
    Exit Sub
End If
sectionCount = ActiveDocument.Sections.Count
If sectionCount = 1 Then
    MsgBox "Only one section found in document, indicating this document has not be styled. Exiting Endnotes macro."
    Exit Sub
End If

' Setup global Endnote settings (continuous number, endnotes at document end, number with integers)
'ActiveDocument.Endnotes.StartingNumber = 1
'ActiveDocument.Endnotes.NumberingRule = wdRestartContinuous
ActiveDocument.Endnotes.NumberingRule = wdRestartSection
ActiveDocument.Endnotes.Location = 1
ActiveDocument.Endnotes.NumberStyle = wdNoteNumberStyleArabic


' See if we're using custom Palgrave tags
iReply = MsgBox("To insert custom Palgrave Endnote tags in your text, click 'YES' " & vbNewLine & vbNewLine & _
    "To continue with standard Endnote ref numbers only: click 'NO'.", vbYesNo + vbExclamation + vbDefaultButton2, "Alert")
If iReply = vbYes Then palgraveTag = True

' Begin working on Endnotes
Application.ScreenUpdating = False

With ActiveDocument
  For Each eNote In .Endnotes
    With eNote
      With .Reference.Characters.First
        .Collapse wdCollapseStart
        BookmarkName = "Endnote" & BookmarkNum
        .Bookmarks.Add Name:=BookmarkName
        .InsertCrossReference wdRefTypeEndnote, wdEndnoteNumberFormatted, eNote.Index
        nref = .Characters.First.Fields(1).Result
        If palgraveTag = False Then
            .Characters.First.Fields(1).Unlink
        Else
            eNote.Reference.InsertBefore "<NoteCallout>" & nref & "</NoteCallout>"   'tags location of old ref
            .Characters.Last.Fields(1).Delete      ' delete old ref
        End If

        'Now for the header business:
        addChapterName = False
        Set refRng = ActiveDocument.Bookmarks(BookmarkName).Range
        refSection = ActiveDocument.Range(0, refRng.Sections(1).Range.End).Sections.Count
        If refSection <> lastRefSection Then
            'following line for debug: comment later
            'MsgBox refSection & " is section of Endnote index #" & nref
            chapterName = endnoteHeader(refSection)
            If chapterName = "```No Header found```" Then
                MsgBox "ERROR: encountered endnote reference in a section lacking any requisite header style (fmh, cn, ct or ctnp)" & vbNewLine & vbNewLine & _
                "Exiting macro, reverting to last save."
                Documents.Open FileName:=ActiveDocument.FullName, Revert:=True
                Application.ScreenUpdating = True
                Exit Sub
            End If
            addChapterName = True
            lastRefSection = refSection
            'following line for debug: comment later
            'MsgBox chapterName
        End If
        BookmarkNum = BookmarkNum + 1
      End With
      .Range.Cut
    End With
    
'''''Since I am not attempting to number at end of each secion,  commenting out parts of this clause
    'If .Range.EndnoteOptions.Location = wdEndOfSection Then
    '  Set nRng = eNote.Range.Sections.First.Range
    'Else
    Set nRng = .Range
    'End If
    With nRng
      .Collapse wdCollapseEnd
      .End = .End - 1
      If .Characters.Last <> Chr(12) Then .InsertAfter vbCr
      If addHeader = True Then
        .InsertAfter "ENDNOTES" & vbCr
        With .Paragraphs.Last.Range
            .Style = "BM Head (bmh)"
        End With
        addHeader = False
      End If
      If addChapterName = True Then
        .InsertAfter chapterName '
        With .Paragraphs.Last.Range
            .Style = "Endnote Text (ntx)"
        End With
      End If
      .InsertAfter nref & " "
      With .Paragraphs.Last.Range
        .Style = "Endnote Text (ntx)"
        .Words.First.Style = "Endnote Reference"
      End With
      .Collapse wdCollapseEnd
      .Paste
      If .Characters.Last = Chr(12) Then .InsertAfter vbCr
    End With
  Next
'''This deletes the endnote
  For Each eNote In .Endnotes
    eNote.Delete
  Next
End With
Set nRng = Nothing

Call RemoveAllBookmarks

Application.ScreenUpdating = True

End Sub

Function endnoteHeader(refSection As Integer) As String
Dim sectionRng As Range
Dim searchStylesArray(4) As String                       ' number of items in array should be declared here
Dim searchTest As Boolean
Dim i As Long

Call zz_clearFindB

Set sectionRng = ActiveDocument.Sections(refSection).Range
searchStylesArray(1) = "FM Head (fmh)"
searchStylesArray(2) = "Chap Number (cn)"
searchStylesArray(3) = "Chap Title (ct)"
searchStylesArray(4) = "Chap Title Nonprinting (ctnp)"
searchTest = False
i = 1

Do Until searchTest = True
Set sectionRng = ActiveDocument.Sections(refSection).Range
With sectionRng.Find
  .ClearFormatting
  .Style = searchStylesArray(i)
  .Wrap = wdFindStop
  .Forward = True
End With
If sectionRng.Find.Execute Then
    endnoteHeader = sectionRng
    searchTest = True
Else
'following line for debug: comment later
    'MsgBox searchStylesArray(i) + " Not Found"
    i = i + 1
    If i = 5 Then
        searchTest = True
        endnoteHeader = "```No Header found```"
    End If
End If
Loop
    
Call zz_clearFindB
    
End Function

Sub RemoveAllBookmarks()

'three options from http://wordribbon.tips.net/T009004_Removing_All_Bookmarks.html
'Version 1
Dim objBookmark As Bookmark
For Each objBookmark In ActiveDocument.Bookmarks
    objBookmark.Delete
Next

'Version 2
'Dim stBookmark As Bookmark
'ActiveDocument.Bookmarks.ShowHidden = True
'If ActiveDocument.Bookmarks.Count >= 1 Then
'   For Each stBookmark In ActiveDocument.Bookmarks
'      stBookmark.Delete
'   Next stBookmark
'End If

'Version 3
'Dim objBookmark As Bookmark
'
'For Each objBookmark In ActiveDocument.Bookmarks
'    If Left(objBookmark.Name, 1) <> "_" Then objBookmark.Delete
'Next


'http://wordribbon.tips.net/T009004_Removing_All_Bookmarks.html

End Sub

Private Sub zz_clearFindB()

Dim clearRng As Range
Set clearRng = ActiveDocument.Words.First

With clearRng.Find
  .ClearFormatting
  .Replacement.ClearFormatting
  .Text = ""
  .Replacement.Text = ""
  .Wrap = wdFindStop
  .Format = False
  .MatchCase = False
  .MatchWholeWord = False
  .MatchWildcards = False
  .MatchSoundsLike = False
  .MatchAllWordForms = False
  .Execute
End With
End Sub


