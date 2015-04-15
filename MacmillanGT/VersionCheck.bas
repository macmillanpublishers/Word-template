Attribute VB_Name = "VersionCheck"
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

' ---------------------------------------------------------
' If re-creating to check another template, change all the variables here
templateFile = "MacmillanGT.dotm"                         'the template file you are checking
pcDir = Environ("APPDATA") & "\Microsoft\Word\STARTUP\"   'the PC directory where templateFile is supposed to live, including trailing slash
macDir = "Macintosh HD:Applications:Microsoft Office 2011:Office:Startup:Word:"  'the Mac directory where templateFile is supposed to live, including trailing colon
''---------------------------------------------------------

'these variables stay the same even if checking a different template
pcTemplatePath = pcDir & templateFile
macTemplatePath = macDir & templateFile
TheOS = System.OperatingSystem

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
Dim macUser As String

TheOS = System.OperatingSystem

'For files located in user directory on Mac. Gives error on PC w/o if-then
If TheOS Like "*Mac*" Then
  macUser = MacScript("tell application " & Chr(34) & "System Events" & Chr(34) & Chr(13) & "return (name of current user)" & Chr(13) & "end tell")
End If

' ---------------------------------------------------------
' If re-creating to check another template, change all the variables here
templateFile = "macmillan.dotm"  'the template file you are checking
pcDir = Environ("PROGRAMDATA") & "\MacmillanStyleTemplate\"  'the directory where templateFile is supposed to live, including training slash
macDir = "Macintosh HD:Users:" & macUser & ":Documents:MacmillanStyleTemplate:"           'the directory where templateFile is supposed to live, including trailing colon
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
Private Sub VersionCheck(fullPath As String, fileName As String)

'------------------------------
'created by Erica Warren 2014-04-08         erica.warren@macmillan.com
'Alerts user to the version number of the template file

Dim installedVersion As String
Debug.Print fullPath

If FileOrDirExists(fullPath) = False Then           ' the template file is not installed, or is not in the correct place
    installedVersion = "none"
Else                                                                'the template file is installed in the correct place
    Documents.Open fileName:=fullPath, ReadOnly:=True                   ' Note can't set Visible:=False because that's not an argument in Word Mac VBA :(
    installedVersion = Documents(fullPath).CustomDocumentProperties("version")
    Documents(fullPath).Close
End If

'Now we tell the user what version they have
If installedVersionST <> "none" Then
    MsgBox "You currently have version " & installedVersion & " of the file " & fileName & " installed."
Else
    MsgBox "You do not have " & fileName & " installed in the right place on your computer."
End If

End Sub
Function FileOrDirExists(PathName As String) As Boolean
     ' From here: http://www.vbaexpress.com/kb/getarticle.php?kb_id=559
     'Macro Purpose: Function returns TRUE if the specified file
     '               or folder exists, false if not.
     'PathName     : Supports Windows mapped drives or UNC
     '             : Supports Macintosh paths
     'File usage   : Provide full file path and extension
     'Folder usage : Provide full folder path
     '               Accepts with/without trailing "\" (Windows)
     '               Accepts with/without trailing ":" (Macintosh)
     
    Dim iTemp As Integer
     
     'Ignore errors to allow for error evaluation
    On Error Resume Next
    iTemp = GetAttr(PathName)
     
     'Check if error exists and set response appropriately
    Select Case Err.Number
    Case Is = 0
        FileOrDirExists = True
    Case Else
        FileOrDirExists = False
    End Select
     
     'Resume error checking
    On Error GoTo 0
End Function

