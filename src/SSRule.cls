VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SSRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstr_rule_name As String
Private mstr_ss_name As String
Private mbln_section_required As Boolean
Private mcol_insert_before As Collection
Private mbln_multiple As Boolean
Private mcol_styles As Collection
Private mstr_position As String
Private mcol_optional_heading_styles As Collection
Private mbln_first_child As Boolean
Private mbln_first_child_text As Collection
Private mbln_first_child_match As Boolean
Private mbln_previous_sibling As Boolean
Private mbln_previous_until As Boolean
Private mcol_required_styles As Collection
Private mcol_previous_until_stop As Collection
Private mbln_last As Boolean


Public Sub Init(Name As String, values As Dictionary, rulenum As Long, section_types As Dictionary)
' To create objects, from calls from the "factory"
' It uses the 'rulenum' as an index, to check against contiguous criteria blocks
' If it's not the last block criteria, it creates a new object with the next block criteria
    
    Dim str_criteria_count As String
    Dim str_next_criteria_count As String
    Dim objNewSSrule As SSRule
    
    str_criteria_count = GetTwoDigitString(rulenum)
    str_next_criteria_count = GetTwoDigitString(rulenum + 1)
    
    mstr_rule_name = Name & "_" & str_criteria_count
    mstr_ss_name = Name
    If values.Exists("section_required") Then
      mbln_section_required = True
      Set mcol_insert_before = values("section_required").Item("insert_before")
    Else
      mbln_section_required = False
      Set mcol_insert_before = New Collection
    End If
    mbln_multiple = CBool(values("contiguous_block_criteria_" & str_criteria_count).Item("multiple"))
    Set mcol_styles = values("contiguous_block_criteria_" & str_criteria_count).Item("styles")
    If values("contiguous_block_criteria_" & str_criteria_count).Exists("position") Then
        mstr_position = values("contiguous_block_criteria_" & str_criteria_count).Item("position")
    Else
        mstr_position = ""
    End If
    If values("contiguous_block_criteria_" & str_criteria_count).Exists("optional_heading_styles") Then
        Set mcol_optional_heading_styles = values("contiguous_block_criteria_" & str_criteria_count).Item("optional_heading_styles")
    Else
        Set mcol_optional_heading_styles = New Collection
    End If
    If values("contiguous_block_criteria_" & str_criteria_count).Exists("first_child") Then
      mbln_first_child = True
      Set mbln_first_child_text = values("contiguous_block_criteria_" & str_criteria_count).Item("first_child").Item("text")
      mbln_first_child_match = CBool(values("contiguous_block_criteria_" & str_criteria_count).Item("first_child").Item("match"))
    Else
      mbln_first_child = False
      Set mbln_first_child_text = New Collection
      mbln_first_child_match = True
    End If
    If values("contiguous_block_criteria_" & str_criteria_count).Exists("previous_sibling") Then
      mbln_previous_sibling = True
      mbln_previous_until = False
      Set mcol_required_styles = values("contiguous_block_criteria_" & str_criteria_count).Item("previous_sibling").Item("required_styles")
      Set mcol_previous_until_stop = New Collection
    Else
      mbln_previous_sibling = False
      mbln_previous_until = True
      Set mcol_required_styles = values("contiguous_block_criteria_" & str_criteria_count).Item("previous_until").Item("required_styles")
      Set mcol_previous_until_stop = values("contiguous_block_criteria_" & str_criteria_count).Item("previous_until").Item("previous_until_stop")
    End If
    ' "mbln_last" value is to let us know if there are more rules coming for this SS; important when 'section_required' = true
    If values.Exists("contiguous_block_criteria_" & str_next_criteria_count) Then
        mbln_last = False
    Else
        mbln_last = True
    End If
    
    ' process this rule!
    Call ProcessRule(section_types)

    ' if there are more contiguous block criteria, create a new object on the spot (via factory)
    If Not mbln_last Then
        Set objNewSSrule = Factory.CreateSSrule(Name, values, rulenum + 1, section_types)
    End If

    
End Sub
    
Private Function GetTwoDigitString(num As Long)
    If num < 9 Then
        GetTwoDigitString = "0" & num
    Else
        GetTwoDigitString = CStr(num)
    End If
End Function


Private Sub ProcessRule(section_types As Dictionary)
'Here we would take our rules object and process content conversions in the doc..?
    
    '' Output test (can delete) :
    Debug.Print "mstr_rule_name : " & mstr_rule_name
    Debug.Print "mstr_ss_name : " & mstr_ss_name
    Debug.Print "mbln_section_required : " & mbln_section_required
    Debug.Print "mcol_insert_before : " & CheckCollection(mcol_insert_before)
    Debug.Print "mbln_multiple : " & mbln_multiple
    Debug.Print "mcol_styles : " & CheckCollection(mcol_styles)
    Debug.Print "mstr_position : " & mstr_position
    Debug.Print "mcol_optional_heading_styles : " & CheckCollection(mcol_optional_heading_styles)
    Debug.Print "mbln_first_child : " & mbln_first_child
    Debug.Print "mbln_first_child_text : " & CheckCollection(mbln_first_child_text)
    Debug.Print "mbln_first_child_match : " & mbln_first_child_match
    Debug.Print "mbln_previous_sibling : " & mbln_previous_sibling
    Debug.Print "mbln_previous_until : " & mbln_previous_until
    Debug.Print "mcol_required_styles : " & CheckCollection(mcol_required_styles)
    Debug.Print "mcol_previous_until_stop : " & CheckCollection(mcol_previous_until_stop)
    Debug.Print "mbln_last : " & mbln_last
    Debug.Print "section_types-FM: " & CheckCollection(section_types("frontmatter"))
    Debug.Print "section_types-main: " & CheckCollection(section_types("main"))
    Debug.Print "section_types-BM: " & CheckCollection(section_types("backmatter"))

End Sub

Private Function CheckCollection(c As Collection)
' assisting in output test in ProcessRule!

    If c.Count > 0 Then
    CheckCollection = c(1)
    Else
    CheckCollection = "empty collection"
    End If
    
End Function
