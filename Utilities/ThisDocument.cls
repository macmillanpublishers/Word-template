VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' by Erica Warren - erica.warren@macmillan.com
' with lots of stuff from here: http://www.cpearson.com/excel/vbe.aspx

' ====== USE ==================================================
' For help with VBA development
' Exports all modules in open templates to the local Word-template git repo
' Shared modules go into "SharedMacros" directory, the rest are
' saved in the same directory as the template they live in;
' also saves the template itself (so if you are working out of Startup,
' it saves a copy in the Word-template/TemplateName repo)

' ===== DEPENDENCIES ==========================================
' Obviously clone the git repo and add its path below
' Each template gets its own subdirectory, name matches exactly (w/o extension)
' Modules that are shared among all template must have name start with "Shared"

' ====== WARNING ==============================================
' from http://www.cpearson.com/excel/vbe.aspx
' Many VBA-based computer viruses propagate themselves by creating and/or modifying
' VBA code. Therefore, many virus scanners may automatically and without warning or
' confirmation delete modules that reference the VBProject object, causing a permanent
' and irretrievable loss of code. Consult the documentation for your anti-virus
' software for details.
'
' So be sure to export and commit often!

Sub ExportAllModules()
    ' Exports all VBA modules in all open template to local git repo
    
    ' ===== ADD PATH TO LOCAL GIT REPO HERE ========================
    ' includes trailing slash
    
    Dim strRepoPath As String
    strRepoPath = "C:\Users\erica.warren\Word-template\"
    
    ' ==============================================================
    
    ' Cycle through each open document
    Dim oDoc As Document
    Dim strExtension As String
    Dim oProject As VBIDE.VBProject
    Dim oModule As VBIDE.VBComponent
    Dim strSharedModules As String
    Dim strDirName As String
    Dim strTemplateModules As String
    
    ' This is where all shared modules go
    strSharedModules = strRepoPath & "SharedModules"
    
    For Each oDoc In Documents
        ' Separate the name and the extension of the document
        strExtension = Right(oDoc.Name, Len(oDoc.Name) - (InStrRev(oDoc.Name, ".") - 1))
        strDirName = Left(oDoc.Name, InStrRev(oDoc.Name, ".") - 1)
        'Debug.Print "File name is " & oDoc.Name
        'Debug.Print "Extension is " & strExtension
        'Debug.Print "Directory is " & strDirName
        
        ' We just want to work with .dotm and .docm (others can't have macros)
        If strExtension = ".dotm" Or strExtension = ".docm" Then
            ' Make sure we're referencing the correct project
            Set oProject = oDoc.VBProject
        
            strTemplateModules = strRepoPath & strDirName
            
            ' Cycle through each module
            For Each oModule In oProject.VBComponents
                ' Select save location based on module name
                If oModule.Name Like "Shared*" Then
                    Call ExportVBComponent(VBComp:=oModule, FolderName:=strSharedModules)
                Else
                    Call ExportVBComponent(VBComp:=oModule, FolderName:=strTemplateModules)
                End If
            Next
            
            ' Might as well save the template file too
            oDoc.Save
            
'            ' ===============================================================================
'            ' This whole bit below is to save the template too, but it's not working yet
'            ' getting errors at Set oDocCopy = Application.Documents.Add(strOldPath, NewTemplate:=True)
'            ' saying "can't execute in break mode", making me think it's trying to execute the
'            ' AutoOpen macro in the copied file, but WordBasic.DisableAutoMacros should have
'            ' solved that so I'm not sure what's going on.
's
'            Dim strOldPath As String
'            Dim strNewPath As String
'            strOldPath = oDoc.Path & Application.PathSeparator & oDoc.Name
'            Debug.Print "Original path is" & strOldPath
'            strNewPath = strTemplateModules & Application.PathSeparator & oDoc.Name
'            Debug.Print "New path is " & strNewPath
'
'            ' Save it where is it open from
'            oDoc.Save
'
'            ' Now, if it's open from somewhere other than the repo, we also want to save
'            ' a copy in the repo but keep this one open to work on
'            ' This is most useful when working on a template in the STARTUP dir, which you may
'            ' do because if you have a template in the STARTUP dir and then open a copy from a different location
'            ' they both show up in the Project Viewer
'            If strNewPath <> strOldPath Then
'                ' If the original template is in the STARTUP dir, we need to diable it because
'                ' we can't create a new doc with the same name as an open doc
'                If InStr(1, LCase(strOldPath), LCase("startup"), vbTextCompare) > 0 Then         'LCase because "startup" was staying in all caps for some reason, UCase wasn't working
'                On Error Resume Next                            'Error = add-in not available, don't need to uninstall
'                    AddIns(strOldPath).Installed = False
'                On Error GoTo 0
'                End If
'
'                ' Now we have to create a copy of the document to save; if we just SaveAs then the open doc
'                ' becomes the one in the new saved location, but we still want it to be open from STARTUP
'                Dim oDocCopy As Document
'                ' Oh also if you create a copy and it has an AutoOpen macro (which MacmillanGT.dotm does)
'                ' it will try to start that macro and cause an error, so here we disable that with = 1 here
'                WordBasic.DisableAutoMacros (1)
'                ' Need NewTemplate = True to carry all code into new doc
'                Set oDocCopy = Application.Documents.Add(strOldPath, NewTemplate:=True)
'                oDocCopy.SaveAs2 FileName:=strNewPath
'                oDocCopy.Close
'
'                ' Now we need to enable the template again if it is in the STARTUP dir
'                If InStr(1, LCase(strOldPath), LCase("startup"), vbTextCompare) > 0 Then         'LCase because "startup" was staying in all caps for some reason, UCase wasn't working
'                On Error Resume Next                            'Error = add-in not available, don't need to uninstall
'                    AddIns(strOldPath).Installed = True
'                On Error GoTo 0
'                End If
'
'                ' Enable AutoOpen macros going forward here with = 0
'                WordBasic.DisableAutoMacros (0)
'
'
'            End If
'            Debug.Print oDoc.Path
            
            
        End If
    Next oDoc
End Sub


Private Sub ExportVBComponent(VBComp As VBIDE.VBComponent, _
                FolderName As String, _
                Optional FileName As String, _
                Optional OverwriteExisting As Boolean = True)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' COPIED FROM http://www.cpearson.com/excel/vbe.aspx
    ' This function exports the code module of a VBComponent to a text
    ' file. If FileName is missing, the code will be exported to
    ' a file with the same name as the VBComponent followed by the
    ' appropriate extension.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Extension As String
    Dim FName As String
    Extension = GetFileExtension(VBComp:=VBComp)
    If Trim(FileName) = vbNullString Then
        FName = VBComp.Name & Extension
    Else
        FName = FileName
        If InStr(1, FName, ".", vbBinaryCompare) = 0 Then
            FName = FName & Extension
        End If
    End If
    
    If StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
        FName = FolderName & FName
    Else
        FName = FolderName & "\" & FName
    End If
    
    If Dir(FName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
        If OverwriteExisting = True Then
            Kill FName
        Else
            Exit Sub
        End If
    End If
    
    VBComp.Export FileName:=FName
    
    End Sub
    
Private Function GetFileExtension(VBComp As VBIDE.VBComponent) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' COPIED FROM http://www.cpearson.com/excel/vbe.aspx
' This returns the appropriate file extension based on the Type of
' the VBComponent.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case VBComp.Type
        Case vbext_ct_ClassModule
            GetFileExtension = ".cls"
        Case vbext_ct_Document
            GetFileExtension = ".cls"
        Case vbext_ct_MSForm
            GetFileExtension = ".frm"
        Case vbext_ct_StdModule
            GetFileExtension = ".bas"
        Case Else
            GetFileExtension = ".bas"
    End Select
    
End Function

