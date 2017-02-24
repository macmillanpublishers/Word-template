Attribute VB_Name = "UnitTestHelpers"
Option Explicit

Function LoadTestDoc(p_strTemplateNameSuffix As String) As Document

Dim strTimestamp As String
Dim strParentFolderPath As String
Dim strTemplatePath As String
Dim strFilePath As String
Dim strFileNameBase As String
Dim strDocxName As String
Dim docTemplate As Document

'set vars
strTimestamp = Format(Now(), "_MM-dd_hh-mm-ss")
strFileNameBase = "Test_MS_"
strDocxName = strFileNameBase & p_strTemplateNameSuffix & strTimestamp & ".docx"

' get parent folder
strParentFolderPath = CreateObject("Scripting.FileSystemObject").Getfile(ThisDocument.FullName).ParentFolder.ParentFolder.Path

' set file paths
strTemplatePath = strParentFolderPath & Application.PathSeparator & "TestDocuments" _
& Application.PathSeparator & strFileNameBase & p_strTemplateNameSuffix & ".dotx"
strFilePath = strParentFolderPath & Application.PathSeparator & "TestDocuments" _
& Application.PathSeparator & strDocxName

' open template, save as new testdoc
Set docTemplate = Documents.Open(fileName:=strTemplatePath, ReadOnly:=True)
docTemplate.SaveAs fileName:=strFilePath

'return the test document
Set LoadTestDoc = Documents(strDocxName)

End Function

Function DeleteTestDoc(p_strTestfilePath As String)

    Documents(p_strTestfilePath).Close SaveChanges:=wdDoNotSaveChanges
    Kill p_strTestfilePath

End Function

