Attribute VB_Name = "SStart_Rules"
Option Explicit

Sub SSRulesBoss()
Dim strJsonFilepath As String
Dim strJsonFilename As String
Dim strJson As String
Dim dictParsedJSON As Dictionary
Dim dictSingleSS As Dictionary
Dim dictSectionTypes As Dictionary
Dim i As Long
Dim strSSname As String
Dim objNewSSrule As SSRule

strJsonFilename = "ss_rules.json"
strJsonFilepath = CreateObject("Scripting.FileSystemObject").Getfile(ThisDocument.FullName).ParentFolder.Path _
& Application.PathSeparator & "bookmaker_validator" & Application.PathSeparator & strJsonFilename

strJson = Utils.ReadTextFile(Path:=strJsonFilepath, FirstLineOnly:=False)
Set dictParsedJSON = JsonConverter.ParseJson(strJson)

Set dictSectionTypes = getSectionTypes(dictParsedJSON)

For i = 0 To dictParsedJSON.Count - 1
    Set dictSingleSS = New Dictionary
    strSSname = dictParsedJSON.Keys(i)
    Set dictSingleSS = dictParsedJSON(strSSname)
    Set objNewSSrule = Factory.CreateSSrule(strSSname, dictSingleSS, 1, dictSectionTypes)
Next

End Sub
Function getSectionTypes(p_dictParsedJSON As Dictionary) As Dictionary

Dim dictSectionTypes As Dictionary
Dim collFrontmatter As Collection
Dim collMain As Collection
Dim collBackmatter As Collection
Dim j As Long
Dim strSSname As String

Set collFrontmatter = New Collection
Set collMain = New Collection
Set collBackmatter = New Collection
Set dictSectionTypes = New Dictionary

For j = 0 To p_dictParsedJSON.Count - 1
    strSSname = p_dictParsedJSON.Keys(j)
    If p_dictParsedJSON(strSSname).Item("section_type") = "frontmatter" Then
        collFrontmatter.Add (strSSname)
    ElseIf p_dictParsedJSON(strSSname).Item("section_type") = "main" Then
        collMain.Add (strSSname)
    ElseIf p_dictParsedJSON(strSSname).Item("section_type") = "backmatter" Then
        collBackmatter.Add (strSSname)
    End If
Next

dictSectionTypes.Add "frontmatter", collFrontmatter
dictSectionTypes.Add "main", collMain
dictSectionTypes.Add "backmatter", collBackmatter

Set getSectionTypes = dictSectionTypes
End Function
