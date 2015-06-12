Attribute VB_Name = "GlobalVars"
'This is Part 2 of 2 of the installer. Part 1 must be saved in the ThisDocument module as
'a sub called "Document_Open" in a .docm file.

'The template file must be uploaded as an attachment to https://confluence.macmillan.com/display/PBL/Test


'=== SET THESE VARIABLES TO DOWNLOAD A SPECIFIC FILE ===========================================
Public strTemplateName As String       'What you're calling the file for your users
Public strTemplateFile As String       'Exact name of the file to be downloaded with extension

Public Sub DeclareGlobalVars()

    strTemplateName = "Macmillan Castoff Macro"
    strTemplateFile = "CastoffMacro.dotm"

End Sub
'===============================================================================================
