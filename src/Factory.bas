Attribute VB_Name = "Factory"
Option Explicit

Public Function CreateSSrule(Name As String, values As Dictionary, rulenum As Long, section_types As Dictionary) As SSRule
    Set CreateSSrule = New SSRule
    CreateSSrule.Init Name:=Name, values:=values, rulenum:=rulenum, section_types:=section_types
End Function
