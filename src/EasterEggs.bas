Attribute VB_Name = "EasterEggs"
Option Explicit

Sub Triceratops()
  Set activeDoc = ActiveDocument
  
    Dim strMessage As String
    Dim strTriceratops As String
    Dim strFindTrike As String
    
    strFindTrike = _
        "                            __.--'~~~~~`--." & Chr(13) & _
        "         ..       __.    .-~               ~-." & Chr(13) & _
        "         ((\     /   `}.~                     `." & Chr(13) & _
        "          \\\  .{     }               /     \   \" & Chr(13) & _
        "      (\   \\~~       }              |       }   \" & Chr(13) & _
        "       \`.-~ -"
    
    strTriceratops = _
        "                            __.--'~~~~~`--." & Chr(13) & _
        "         ..       __.    .-~               ~-." & Chr(13) & _
        "         ((\     /   `}.~                     `." & Chr(13) & _
        "          \\\  .{     }               /     \   \" & Chr(13) & _
        "      (\   \\~~       }              |       }   \" & Chr(13) & _
        "       \`.-~ -@~     }  ,-,.         |       )    \" & Chr(13) & _
        "       (___/     ) _}  (    :        |    / /      `._" & Chr(13) & _
        "        `----._-~.     _\ \ |_       \   / /-.__     `._" & Chr(13) & _
        "               ~~----~~  \ \| ~~--~~~(  + /     ~-._    ~-._" & Chr(13) & _
        "                         /  /         \  \          ~--.,___~_-_." & Chr(13) & _
        "                      __/  /          _\  )" & Chr(13) & _
        "                    .<___.'         .<___/`"
    
    ' Search for first 250 characters of triceratops (max search = 255 char)
    With Selection.Find
        .ClearFormatting
        .Text = strFindTrike
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = False
        .MatchCase = False
    End With
    
    If Selection.Find.Execute = True Then       ' There is already one in the doc
        'Expand the selection one backtick at a time until whole thing is selected.
        Selection.Extend ("`")
        Selection.Extend ("`")
        Selection.Extend ("`")
        Selection.Extend ("`")
        If Selection.Text = strTriceratops Then     ' Make sure we've got what we want
            strMessage = "You appear to have a Triceratops loose in your document. Would you like me to remove it?"
            If MsgBox(strMessage, vbYesNo, "Extinct Wildlife Alert") = vbNo Then
                Exit Sub
            Else
                Selection.Delete
            End If
        Else
            ' selection doesn't match Triceratops string
            strMessage = "We found something, but it's not quite a Triceratops."
            MsgBox strMessage, vbCritical, "What's This?"
        End If
    Else    ' There is NOT a triceratops in the document yet. Let's add one!
        strMessage = "Would you like a Triceratops at the end of your document?"
        If MsgBox(strMessage, vbYesNo, "Dinosaur Not Found") = vbNo Then
            Exit Sub
        Else
            'Turn off nonprinting characters so spaces don't show
            activeDoc.ActiveWindow.View.ShowAll = False
            
            With Selection
                ' move to end of document (so we don't mess up anyone's text
                .EndKey Unit:=wdStory
                ' insert a new paragraph (ditto)
                .InsertAfter (Chr(13))
                ' move to that new paragraph
                .EndKey Unit:=wdStory
                ' Must be monospace font to see the Triceratops
                .Font.Name = "Courier New"
                ' 10 pt so it fits the width of the page
                .Font.Size = 10
                    ' All this for best view of Triceratops
                    With .ParagraphFormat
                    .Alignment = wdAlignParagraphLeft
                    .FirstLineIndent = 0
                    .LineSpacingRule = wdLineSpaceSingle
                    .SpaceAfter = 0
                    .SpaceBefore = 0
                    End With
            End With
                ' Add the Triceratops!
                Selection.InsertAfter Text:=strTriceratops
                Selection.Collapse (wdCollapseEnd)
        End If
    End If
End Sub
