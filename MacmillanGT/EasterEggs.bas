Attribute VB_Name = "EasterEggs"
Option Explicit

Sub Welcome()

Debug.Print Weekday(Date)
Dim strUsername As String
Dim strMessage As String

'This is the username used for Track Changes. How does this get entered by default?
strUsername = Application.UserName

'Weekday() returns an integer for the day of the week, starting at 1 on Sunday.
If Weekday(Date) = 2 Then       'It's Monday!
    strMessage = "Hello, " & strUsername & "!"
    MsgBox strMessage, vbDefaultButton1, "Ahoy!"
End If
End Sub

Sub Triceratops()
    Dim strMessage As String
    strMessage = "Would you like a Triceratops at the end of your document?"
    If MsgBox(strMessage, vbOKCancel) = vbCancel Then
        Exit Sub
    Else
        Selection.HomeKey Unit:=wdStory
        Selection.InsertAfter (Chr(13))
        Selection.HomeKey Unit:=wdStory

        With Selection.ParagraphFormat
            .Alignment = wdAlignParagraphLeft
            .FirstLineIndent = 0
            .LineSpacing = 1
            .SpaceAfter = 0
            .SpaceBefore = 0
            .
        
End Sub
