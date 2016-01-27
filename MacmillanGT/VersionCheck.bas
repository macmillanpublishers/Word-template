Attribute VB_Name = "VersionCheck"
Option Explicit
Sub CheckMacmillanGT()

'----------------------------------
    'created by Erica Warren 2014-04-08     erica.warren@macmillan.com
    'Creates a toolbar button that tells the user the current version of the installed template when pressed.
    '----------------------------------
    
    Dim pcDir As String
    Dim macDir As String
    Dim templateFile As String
    Dim pcTemplatePath As String
    Dim macTemplatePath As String
    Dim TheOS As String
    TheOS = System.OperatingSystem
    Dim strMacDocs As String
    
    ' ---------------------------------------------------------
    ' If re-creating to check another template, change all the variables here
    templateFile = "MacmillanGT.dotm"  'the template file you are checking
    pcDir = Environ("PROGRAMDATA") & "\MacmillanStyleTemplate\"  'the directory where templateFile is supposed to live, including training slash
    strMacDocs = MacScript("return (path to documents folder) as string")
    macDir = strMacDocs & "MacmillanStyleTemplate:"
    
    'these variables stay the same even if checking a different template
    pcTemplatePath = pcDir & templateFile
    macTemplatePath = macDir & templateFile

    
    'Pass arguments to VersionCheck sub based on OS
    If Not TheOS Like "*Mac*" Then                  'I am Windows
        Call VersionCheck(pcTemplatePath, templateFile)
    Else                                                            ' I am Mac
        Call VersionCheck(macTemplatePath, templateFile)
    End If

End Sub
Sub CheckMacmillan()

    '----------------------------------
    'created by Erica Warren 2014-04-08     erica.warren@macmillan.com
    'Creates a toolbar button that tells the user the current version of the installed template when pressed.
    '----------------------------------
    
    Dim pcDir As String
    Dim macDir As String
    Dim templateFile As String
    Dim pcTemplatePath As String
    Dim macTemplatePath As String
    Dim TheOS As String
    Dim strMacDocs
    
    TheOS = System.OperatingSystem
    
    
    ' ---------------------------------------------------------
    ' If re-creating to check another template, change all the variables here
    templateFile = "macmillan.dotm"  'the template file you are checking
    pcDir = Environ("PROGRAMDATA") & "\MacmillanStyleTemplate\"  'the directory where templateFile is supposed to live, including training slash
    strMacDocs = MacScript("return (path to documents folder) as string")
    macDir = strMacDocs & "MacmillanStyleTemplate:"
    ''---------------------------------------------------------
    
    'these variables stay the same even if checking a different template
    pcTemplatePath = pcDir & templateFile
    macTemplatePath = macDir & templateFile
    
    
    'Pass arguments to VersionCheck sub based on OS
    If Not TheOS Like "*Mac*" Then                  'I am Windows
        Call VersionCheck(pcTemplatePath, templateFile)
    Else                                                            ' I am Mac
        Call VersionCheck(macTemplatePath, templateFile)
    End If

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
