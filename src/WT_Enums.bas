Attribute VB_Name = "WT_Enums"
Option Explicit

' =============================================================================
'       Module to keep all public enumerations in one place
' =============================================================================
' Also so we can create functions to convert strings to enums?

' ===== SectionType ===========================================================
' For top-level book parts: frontmatter, main, backmatter.
' Using bitwise values in case we need to handle things like Acks, which can be
' either frontmatter or backmatter.

Public Enum WT_SectionType
  no_section = 0
  frontmatter = 2 ^ 0
  main = 2 ^ 1
  backmatter = 2 ^ 2
End Enum


' =============================================================================
'   Convert from string to enum
' =============================================================================
' Native functions to convert to a type are CStr(), CBool(), CDate(), etc. so
' use same syntax.

Public Function CSectionType(ConvertFrom As String) As WT_SectionType
  Select Case ConvertFrom
    Case "frontmatter"
      CSectionType = WT_SectionType.frontmatter
    Case "main"
      CSectionType = WT_SectionType.main
    Case "backmatter"
      CSectionType = WT_SectionType.backmatter
  End Select
End Function

