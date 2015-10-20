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
            Next oModule
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

