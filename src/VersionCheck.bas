Attribute VB_Name = "VersionCheck"
Option Explicit
Sub CheckMacmillanGT()
' can't change name to Word-template because "CheckMacmillanGT" is in customUI.xml
' and I don't want to muck with it right now.
'----------------------------------
    'created by Erica Warren 2014-04-08     erica.warren@macmillan.com
    'Creates a toolbar button that tells the user the current version of the installed template when pressed.
    '----------------------------------
    
    Dim templateFile As String
    Dim strMacDocs As String
    Dim strTemplatePath As String

    
    templateFile = "Word-template.dotm"  'the template file you are checking
    strTemplatePath = SharedFileInstaller.StyleDir()

    Call VersionCheck(strTemplatePath, templateFile)


End Sub
Sub CheckMacmillan()

    '----------------------------------
    'created by Erica Warren 2014-04-08     erica.warren@macmillan.com
    'Creates a toolbar button that tells the user the current version of the installed template when pressed.
    '----------------------------------
    
    Dim templateFile As String
    Dim strTemplatePath As String
    
    templateFile = "macmillan.dotx"  'the template file you are checking
    strTemplatePath = SharedFileInstaller.StyleDir()
    
    Call VersionCheck(strTemplatePath, templateFile)

End Sub
Private Sub VersionCheck(fullPath As String, FileName As String)

    '------------------------------
    'created by Erica Warren 2014-04-08         erica.warren@macmillan.com
    'Alerts user to the version number of the template file
    
    Dim installedVersion As String
    'DebugPrint fullPath
    
    If IsItThere(fullPath) = False Then           ' the template file is not installed, or is not in the correct place
        installedVersion = "none"
    Else                                                                'the template file is installed in the correct place
        Documents.Open FileName:=fullPath, ReadOnly:=True                   ' Note can't set Visible:=False because that's not an argument in Word Mac VBA :(
        installedVersion = Documents(fullPath).CustomDocumentProperties("version")
        Documents(fullPath).Close
    End If
    
    'Now we tell the user what version they have
    If installedVersion <> "none" Then
        MsgBox "You currently have version " & installedVersion & " of the file " & FileName & " installed."
    Else
        MsgBox "You do not have " & FileName & " installed on your computer."
    End If

End Sub
