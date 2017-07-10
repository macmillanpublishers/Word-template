Attribute VB_Name = "AttachTemplateMacro"
Option Explicit
'created by Erica Warren - erica.warren@macmillan.com
' ======== PURPOSE =================
' Attaches assorted templates with custom Macmillan styles to the current document

' ======== DEPENDENCIES ============
' 1. Requires MacroHelpers module be installed in the same template
' 2. Requires the macmillan style templates be saved in the correct directories
'    that were nstalled on user's computer with Installer file or updated from Word-template.dotm

''' CHECK IT OUT

Sub zz_AttachStyleTemplate()
    Call AttachMe("macmillan.dotx")
End Sub

Sub zz_AttachBoundMSTemplate()
    Call AttachMe("macmillan_NoColor.dotx")
End Sub

Sub zz_AttachCoverTemplate()
    Call AttachMe("macmillan_CoverCopy.dotm")
End Sub

Sub AttachMe(TemplateName As String)
'Attaches a style template from the MacmillanStyleTemplate directory

' Get version number of style template to add to doc properties
  Dim strVersionFileName As String
  strVersionFileName = Utils.GetFileNameOnly(TemplateName) & ".txt"

  Dim dictVersionInfo As Dictionary
  Set dictVersionInfo = SharedFileInstaller.FileInfo(strVersionFileName)

  Dim strVersionPath As String
  Dim strVersionNumber As String
  
  strVersionPath = dictVersionInfo("Final")

  If Utils.IsItThere(strVersionPath) = False Then
  ' Download version file if we don't have it
    SharedFileInstaller.DownloadFromGithub strVersionFileName
  End If
  strVersionNumber = Utils.ReadTextFile(strVersionPath)

' Get path to actual template
  Dim dictTemplateInfo As Dictionary
  Set dictTemplateInfo = SharedFileInstaller.FileInfo(TemplateName)

  Dim strTemplatePath As String
  strTemplatePath = dictTemplateInfo("Final")

  ' Can't attach template to another template, so
  If IsTemplate(ActiveDocument) = False Then
    'Check that file exists
    If IsItThere(strTemplatePath) = True Then
    
      'Apply template with Styles
      With ActiveDocument
        .UpdateStylesOnOpen = True
        .AttachedTemplate = strTemplatePath
      End With
      SetStyleVersion VersionNumber:=strVersionNumber
    Else
      MsgBox "That style template doesn't seem to exist." & vbNewLine & vbNewLine & _
        "Install the Macmillan Style Template and try again, or contact workflows@macmillan.com for assistance.", _
        vbCritical, "Oh no!"
    End If
  End If
    
End Sub

Private Sub SetStyleVersion(VersionNumber As String)
  Dim strPropName As String
  strPropName = "Version"
  
    If Utils.DocPropExists(objDoc:=ActiveDocument, PropName:=strPropName) Then
        ActiveDocument.CustomDocumentProperties(strPropName).Value = VersionNumber
    Else
        ActiveDocument.CustomDocumentProperties.Add Name:=strPropName, LinkToContent:=False, _
            Type:=msoPropertyTypeString, Value:=VersionNumber
    End If

End Sub

Private Function IsTemplate(ByVal objDoc As Document) As Boolean
  Select Case objDoc.SaveFormat
    Case wdFormatTemplate, _
         wdFormatXMLTemplate, wdFormatXMLTemplateMacroEnabled, _
         wdFormatFlatXMLTemplate, wdFormatFlatXMLTemplateMacroEnabled
      IsTemplate = True
    Case Else
      IsTemplate = False
  End Select
End Function

