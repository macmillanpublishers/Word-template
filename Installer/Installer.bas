Attribute VB_Name = "Installer"
Option Explicit

Sub AutoOpen()
'created by Erica Warren - erica.warren@macmillan.com
'installs the MacmillanGT.dotm template in STARTUP dir
    
    Dim strWelcome As String
    
    strWelcome = "Welcome to the Macmillan Style Template Installer!" & vbNewLine & vbNewLine & _
    "Please click OK to begin the installation. It should take less than a minute."
    
    If MsgBox(strWelcome, vbOKCancel, "Macmillan Style Template") = vbCancel Then
        Exit Sub
    End If
    
    Dim strStartupDir As String
    Dim strTmpDir As String
    Dim strGtFile As String
    Dim strGtTmpPath As String
    Dim strGtFinalPath As String
    Dim blnGtExists As Boolean

    strGtFile = "MacmillanGT.dotm"
    
    strStartupDir = Application.StartupPath
        'Debug.Print "STARTUP: " & strStartupDir
    
    strTmpDir = Environ("TEMP")
        'Debug.Print "TEMP: " & strTmpDir

    strGtTmpPath = strTmpDir & "\" & strGtFile
        'Debug.Print "Full TEMP path: " & strGtTmpPath
    
    strGtFinalPath = strStartupDir & "\" & strGtFile
        'Debug.Print "Full STARTUP path: " & strGtFinalPath
        
    'try to download the Macmillan template from Public Confluence page
    Dim myURL As String
    Dim WinHttpReq As Object
    Dim oStream As Object
    Dim templateURL As String
    
    'this is download link, actual page housing template is http://confluence.macmillan.com/display/PBL/Test
    templateURL = "https://confluence.macmillan.com/download/attachments/9044274/"
    myURL = templateURL & strGtFile

    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False
    
    On Error Resume Next
    WinHttpReq.Send

        ' Exit sub if error in connecting to website
        If Err.Number <> 0 Then 'HTTP request is not OK
            'Debug.Print WinHttpReq.Status
            MsgBox "There is an error trying to download the Macmillan template." & vbNewLine & vbNewLine & _
            "Please check your internet connection or contact workflows@macmillan.com for help."
            Exit Sub
        End If
    On Error GoTo 0

    If WinHttpReq.Status = 200 Then  ' 200 = HTTP request is OK
    
        'if connection OK, download file to temp dir
        myURL = WinHttpReq.responseBody
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile strGtTmpPath, 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
        Set oStream = Nothing
        Set WinHttpReq = Nothing
    End If
    
    'If MacmillanGT already exists in Startup dir, disable it and delete it
    Dim blnNewFile As Boolean
    If Dir(strGtFinalPath) <> vbNullString Then
        blnNewFile = False
        AddIns(strGtFinalPath).Installed = False
        Kill strGtFinalPath
    Else
        blnNewFile = True
    End If
    
    'If delete was successful, move downloaded file to Startup folder
    
    If Dir(strGtFinalPath) = vbNullString Then
        Name strGtTmpPath As strGtFinalPath
    Else
        MsgBox "There was an error installing the Macmillan template." & vbNewLine & _
        "Please close all other Word documents and try again, or contact workflows@macmillan.com."
    End If
    
    'Load new GT file as global template
    If blnNewFile = False Then
        AddIns(strGtFinalPath).Installed = True
    End If
    
    MsgBox "The Macmillan Style Template has been installed on your computer."
    
End Sub
