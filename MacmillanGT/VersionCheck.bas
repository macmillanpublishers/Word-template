Attribute VB_Name = "VersionCheck"
Option Explicit
Sub CheckMacmillanGT()

'----------------------------------
    'created by Erica Warren 2014-04-08     erica.warren@macmillan.com
    'Creates a toolbar button that tells the user the current version of the installed template when pressed.
    '----------------------------------
    
    Dim templateFile As String
    Dim strMacDocs As String
    Dim strTemplatePath As String

    
    templateFile = "MacmillanGT.dotm"  'the template file you are checking
    
    #If Mac Then
        strMacDocs = MacScript("return (path to documents folder) as string")
        strTemplatePath = strMacDocs & "MacmillanStyleTemplate:" & templateFile
    #Else
        strTemplatePath = Environ("PROGRAMDATA") & "\MacmillanStyleTemplate\" & templateFile
    #End If

    
    Call VersionCheck(strTemplatePath, templateFile)


End Sub
Sub CheckMacmillan()

    '----------------------------------
    'created by Erica Warren 2014-04-08     erica.warren@macmillan.com
    'Creates a toolbar button that tells the user the current version of the installed template when pressed.
    '----------------------------------
    
    Dim templateFile As String
    Dim strMacDocs As String
    Dim strTemplatePath As String

    
    templateFile = "macmillan.dotm"  'the template file you are checking
    
    #If Mac Then
        strMacDocs = MacScript("return (path to documents folder) as string")
        strTemplatePath = strMacDocs & "MacmillanStyleTemplate:" & templateFile
    #Else
        strTemplatePath = Environ("PROGRAMDATA") & "\MacmillanStyleTemplate\" & templateFile
    #End If

    
    Call VersionCheck(strTemplatePath, templateFile)

End Sub
Private Sub VersionCheck(fullPath As String, FileName As String)

    '------------------------------
    'created by Erica Warren 2014-04-08         erica.warren@macmillan.com
    'Alerts user to the version number of the template file
    
    Dim installedVersion As String
    'Debug.Print fullPath
    
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
