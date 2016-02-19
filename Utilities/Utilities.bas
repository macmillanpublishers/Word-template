Attribute VB_Name = "Utilities"
Option Explicit
Option Base 1

' ====== ADD PATH TO LOCAL GIT REPO HERE ========================
' DON'T include trailing slash
Private Const strRepoPath = "C:\Users\erica.warren\Word-template"

' by Erica Warren - erica.warren@macmillan.com
' Good advice from here: http://www.cpearson.com/excel/vbe.aspx

' ====== USE ==================================================
' For help with VBA development
' Exports all modules in open templates to the local Word-template git repo (ABOVE!!)
' Shared modules go into "SharedMacros" directory, the rest are
' saved in the same directory as the template they live in

' Also imports all modules saved in git repo to the open templates.
' great for dealing with template file merge conflicts!

' ===== DEPENDENCIES ==========================================
' Obviously clone the git repo and add its path ABOVE
' Each template gets its own subdirectory in the repo, name matches exactly (w/o extension)
' Modules that are shared among all template must have name start with "Shared"

' Not tested on Mac, because saving templates on Mac causes all kinds of nonsense

' ====== WARNING ==============================================
' advice from http://www.cpearson.com/excel/vbe.aspx
' Many VBA-based computer viruses propagate themselves by creating and/or modifying
' VBA code. Therefore, many virus scanners may automatically and without warning or
' confirmation delete modules that reference the VBProject object, causing a permanent
' and irretrievable loss of code. Consult the documentation for your anti-virus
' software for details.
'
' So be sure to export and commit often!

Sub ExportAllModules()
    ' Exports all VBA modules in all open templates to local git repo
    
    ' Cycle through each open document
    Dim oDoc As Document
    Dim strExtension As String
    Dim oProject As VBIDE.VBProject
    Dim oModule As VBIDE.VBComponent
    Dim strSharedModules As String
    Dim strDirName As String
    Dim strTemplateModules As String
    
    ' This is where all shared modules go
    strSharedModules = strRepoPath & Application.PathSeparator & "SharedModules"
    
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
        
            strTemplateModules = strRepoPath & Application.PathSeparator & strDirName
            
            ' Cycle through each module
            For Each oModule In oProject.VBComponents
                ' Select save location based on module name
                If oModule.Name Like "Shared*" Then
                    Call ExportVBComponent(VBComp:=oModule, FolderName:=strSharedModules)
                Else
                    Call ExportVBComponent(VBComp:=oModule, FolderName:=strTemplateModules)
                End If
            Next
            
            ' And also save the template file in the repo if it's not open from there
            ' CopyTemplateToRepo closes and re-opens the doc, so don't use it for THIS doc
            If oDoc.Name <> ThisDocument.Name Then
                CopyTemplateToRepo TemplateDoc:=oDoc
            Else
                'Debug.Print ThisDocument.Name
                oDoc.Save
            End If
            
        End If
    Next oDoc
End Sub


Private Sub ExportVBComponent(VBComp As VBIDE.VBComponent, _
                FolderName As String, _
                Optional FileName As String, _
                Optional OverwriteExisting As Boolean = True)

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
    
    'Debug.Print FName
    
    If Extension = ".frm" Then
        Dim strBinaryFile As String
        
        strBinaryFile = Left(FName, Len(FName) - 1) & "x"
        'Debug.Print strBinaryFile
        
        Dim strShellCmd As String
        strShellCmd = "cmd.exe /C C: & cd " & strRepoPath & " & git checkout " & strBinaryFile
        strShellCmd = Replace(strShellCmd, "\", "\\")
        
        'Debug.Print strShellCmd
        
        Dim result As Variant
        
        result = Shell(strShellCmd, vbMinimizedNoFocus)
        'Debug.Print result
    End If
    
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


Sub ImportAllModules()
    ' Removes all modules in all open template
    ' and reimports them from the local Word-template git repo
    ' SO BE SURE TO EXPORT EVERYTHING BEFORE YOU USE THIS!!
    
    Dim oDocument As Document
    Dim strExtension As String              ' extension of current document
    Dim strSubDirName As String             ' name of subdirectory of template in repo
    Dim strDirInRepo(1 To 2) As String      ' declare number of items in array
    Dim strModuleExt(1 To 3) As String     ' declare number of items in array
    Dim strModuleFileName As String         ' file name with extension, no path
    Dim a As Long
    Dim b As Long
    Dim Counter As Long
    Dim VBComp As VBIDE.VBComponent     ' object for module we're importing
    Dim strFullModulePath As String     ' full path to module with extension
    Dim strModuleName As String         ' Just the module name w/ no extension
    Dim tempVBComp As VBIDE.VBComponent ' Temp module to import ThisDocument code
    Dim currentVBProject As VBIDE.VBProject     ' object of the VB project the modules are in
    Dim strNewCode As String            ' New code in ThisDocument.cls module
    
    For Each oDocument In Documents
        ' We don't want to run this on this code here
        If oDocument.Name <> ThisDocument.Name Then
            strExtension = Right(oDocument.Name, Len(oDocument.Name) - (InStrRev(oDocument.Name, ".") - 1))
            strSubDirName = Left(oDocument.Name, InStrRev(oDocument.Name, ".") - 1)
            'Debug.Print "File name is " & oDocument.Name
            'Debug.Print "Extension is " & strExtension
            'Debug.Print "Directory is " & strSubDirName
            
            ' We just want to work with .dotm and .docm (others can't have macros)
            If strExtension = ".dotm" Or strExtension = ".docm" Then
                ' an array of the directories we're going to be adding modules from
                ' every template gets (1) all modules in its directory and (2) all shared modules
                strDirInRepo(1) = strRepoPath & Application.PathSeparator & strSubDirName & Application.PathSeparator
                strDirInRepo(2) = strRepoPath & Application.PathSeparator & "SharedModules" & Application.PathSeparator
                      
                ' an array of file extensions we're importing, since there are other files in the repo
                strModuleExt(1) = "bas"
                strModuleExt(2) = "cls"
                strModuleExt(3) = "frm"
                
                ' Get rid of all code currently in there, so we don't create duplicates
                Call DeleteAllVBACode(oDocument)
                
                ' set the Project object for this document
                Set currentVBProject = Nothing
                Set currentVBProject = oDocument.VBProject
                
                ' loop through the two directories
                For a = LBound(strDirInRepo()) To UBound(strDirInRepo())
                    ' for each directory, loop through all files of each extension
                    For b = LBound(strModuleExt()) To UBound(strModuleExt())
                        ' with the Dir function this returns just the files in this directory
                        strModuleFileName = Dir(strDirInRepo(a) & "*." & strModuleExt(b))
                        ' so loop through each file of that extension in that directory
                        Do While strModuleFileName <> "" And Counter < 100
                            Counter = Counter + 1               ' to prevent infinite loops
                            'Debug.Print strModuleFileName
                            
                            strModuleName = Left(strModuleFileName, InStrRev(strModuleFileName, ".") - 1)
                            strFullModulePath = strDirInRepo(a) & strModuleFileName
                            'Debug.Print "Full path to module is " & strFullModulePath
                            
                            ' Resume Next because Set VBComp = current project will cause an error if that
                            ' module doesn't exist, and it doesn't because we just deleted everything
                            On Error Resume Next
                            Set VBComp = Nothing
                            Set VBComp = currentVBProject.VBComponents(strModuleName)
                            
                            ' So if that Set VBComp failed because it doesnt' exist, add it!
                            If VBComp Is Nothing Then
                                currentVBProject.VBComponents.Import FileName:=strFullModulePath
                            Else    ' it DOES exist already
                                ' See then if it's the "ThisDocument" module, which can't be deleted
                                ' So we can't import because it would just create a duplicate, not replace
                                If VBComp.Type = vbext_ct_Document Then
                                    ' sp we'll create a temp module of the module we want to import
                                    Set tempVBComp = currentVBProject.VBComponents.Import(strFullModulePath)
                                    ' then delete the content of ThisDocument and replace it with the content
                                    ' of the temp module
                                    With VBComp.CodeModule
                                        .DeleteLines 1, .CountOfLines
                                        strNewCode = tempVBComp.CodeModule.lines(1, tempVBComp.CodeModule.CountOfLines)
                                        .InsertLines 1, strNewCode
                                    End With
                                    On Error GoTo 0
                                    ' then remove the temp module
                                    currentVBProject.VBComponents.Remove tempVBComp
                                End If
                            End If
                            ' have to do this to make the Dir function loop through all files
                            strModuleFileName = Dir()
                        Loop
                        
                        'Debug.Print strModuleFileName
                    Next b
                Next a
                
            End If
        
            ' And then save the updated template in the repo
            CopyTemplateToRepo TemplateDoc:=oDocument
            
        End If

    Next oDocument
    
    
End Sub


Sub DeleteAllVBACode(objTemplate As Document)
    ' Again copied from http://www.cpearson.com/excel/vbe.aspx
    ' Though made it take an argument
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    
    Set VBProj = objTemplate.VBProject
    
    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_Document Then
            Set CodeMod = VBComp.CodeModule
            With CodeMod
                .DeleteLines 1, .CountOfLines
            End With
        Else
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp
End Sub


Sub InsertShapes()
    ' Inserts pictures for use in creating Mac toolbar in MacmillanGT.dotm
    ' But this way we can name them
    
    Dim strPath As String
    Dim strFileName As String
    Dim shpToolbar As Shape
    Dim lngStart As Long
    Dim lngNumChar As Long
    Dim strShapeName As String
    Dim shpNewPicture As Shape
    
    strPath = "C:\Users\erica.warren\Dropbox (Macmillan Publishers)\Icons\16px gray\"
    strFileName = Dir$(strPath & "*.*")
    
    Do While strFileName <> ""
        ' Get shape name
        lngStart = InStr(strFileName, "---") + 3
        lngNumChar = InStr(strFileName, ".") - lngStart
        strShapeName = Mid(strFileName, lngStart, lngNumChar)
        
        Debug.Print strShapeName
        
        ' Insert picture
        Set shpNewPicture = ActiveDocument.Shapes.AddPicture(FileName:=strPath & strFileName)
        shpNewPicture.Name = strShapeName
        
        strFileName = Dir$
    Loop
    
End Sub


Sub CopyTemplateToRepo(TemplateDoc As Document)
' copies the current template file to the local git repo
    
    ' Don't copy if it's already open from the repo
    If strRepoPath <> TemplateDoc.Path Then
        
        Dim strDir As String
        Dim strCurrentTemplatePath As String
        Dim strDestinationFilePath As String
        
        ' Get name of template w/o extension
        strDir = Left(TemplateDoc.Name, InStrRev(TemplateDoc.Name, ".") - 1)
        
        ' Current file full path, to open after copy
        strCurrentTemplatePath = TemplateDoc.FullName
        'Debug.Print strCurrentTemplatePath
        
        ' location in repo
        strDestinationFilePath = strRepoPath & Application.PathSeparator & strDir & _
            Application.PathSeparator & TemplateDoc.Name
        'Debug.Print strDestinationFilePath
        
        ' Template needs to be closed for FileCopy to work
        TemplateDoc.Close SaveChanges:=wdSaveChanges
        
        ' Also not installed as an add-in
        AddIns(strCurrentTemplatePath).Installed = False
        
        FileCopy Source:=strCurrentTemplatePath, Destination:=strDestinationFilePath
        
        ' Reinstall add-in
        ' Should add check if it's a template, not a .docm
        WordBasic.DisableAutoMacros
        AddIns(strCurrentTemplatePath).Installed = True
        
        ' And then open the document again.
        Documents.Open FileName:=strCurrentTemplatePath, _
                        ReadOnly:=False, _
                        Revert:=False
        
    End If
    
End Sub


