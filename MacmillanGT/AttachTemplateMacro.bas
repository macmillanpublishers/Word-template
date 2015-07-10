Attribute VB_Name = "AttachTemplateMacro"
Option Explicit

Sub zz_AttachStyleTemplate()
    Call AttachMe("macmillan.dotm")
End Sub

Sub zz_AttachBoundMSTemplate()
    Call AttachMe("macmillan_NoColor.dotm")
End Sub

Sub zz_AttachCoverTemplate()
    Call AttachMe("MacmillanCoverCopy.dotm")
End Sub

Private Sub AttachMe(TemplateName As String)
'Attaches a style template from the MacmillanStyleTemplate directory

    Dim currentUser As String
    Dim myFile As String
        
    'Set template path according to OS
    #If Mac Then
        currentUser = MacScript("tell application " & Chr(34) & "System Events" & Chr(34) & Chr(13) & _
            "return (name of current user)" & Chr(13) & "end tell")
        myFile = "Macintosh HD:Users:" & currentUser & ":Documents:MacmillanStyleTemplate:" & TemplateName
    #Else
        myFile = Environ("PROGRAMDATA") & "\MacmillanStyleTemplate\" & TemplateName
    #End If
        
        'Check that file exists
        If IsItThere(myFile) = True Then
        
            'Apply template with Styles
            With ActiveDocument
                .UpdateStylesOnOpen = True
                .AttachedTemplate = myFile
            End With
        Else
            MsgBox "That style template doesn't seem to exist." & vbNewLine & vbNewLine & _
                    "Install the Macmillan Style Template and try again, or contact workflows@macmillan.com for assistance.", _
                    vbCritical, "Oh no!"
        End If
        
    End Sub
