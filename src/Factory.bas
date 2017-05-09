Attribute VB_Name = "Factory"
Option Explicit

Public Function CreateSSrule(Name As String, Values As Dictionary, rulenum As Long) As SSRule ' section_types As Dictionary) As SSRule
    Set CreateSSrule = New SSRule
    CreateSSrule.Init Name:=Name, Values:=Values, rulenum:=rulenum ', section_types:=section_types
End Function

