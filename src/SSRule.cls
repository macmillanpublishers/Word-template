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

' *****************************************************************************
'       PRIVATE VARIABLES
' *****************************************************************************

' Private property var declarations
Private prop_rule_name As String
Private prop_ss_name As String
Private prop_section_required As Boolean
Private prop_insert_before As Collection
Private prop_position As String
Private prop_multiple As Boolean
Private prop_styles As Collection
Private prop_optional_heading_styles As Collection
Private prop_first_child As Boolean
Private prop_first_child_text As Collection
Private prop_first_child_match As Boolean
Private prop_required_styles As Collection
Private prop_previous_until As Collection
Private prop_last As Boolean
Private prop_priority As Long

' module/class only var declarations
Private m_dictValues As Dictionary
Private m_name As String
Private m_str_criteria_count As String
Private m_str_next_criteria_count As String

' *****************************************************************************
'       PUBLIC PROPERTIES
' *****************************************************************************
' all of these are Read-Only, and refer to their equivalent values in the SSRules JSON

Public Property Get RuleName() As String
    prop_rule_name = prop_ss_name & "_" & m_str_criteria_count
    RuleName = prop_rule_name
End Property
Public Property Get SectionName() As String
    SectionName = prop_ss_name
End Property
Public Property Get SectionRequired() As Boolean
    If m_dictValues.Exists("section_required") Then
        If m_dictValues("section_required").Item("value") = True Then
            prop_section_required = True
        Else
            prop_section_required = False
        End If
    Else
      prop_section_required = False
    End If
    SectionRequired = prop_section_required
End Property
Public Property Get InsertBefore() As Collection
    If m_dictValues.Exists("section_required") Then
      Set prop_insert_before = m_dictValues("section_required").Item("insert_before")
    Else
      Set prop_insert_before = New Collection
    End If
    InsertBefore = prop_insert_before
End Property
Public Property Get Position() As String
    If m_dictValues.Exists("position") Then
      prop_position = m_dictValues("position")
    Else
      prop_position = ""
    End If
    Position = prop_position
End Property
Public Property Get Multiple() As Boolean
    prop_multiple = CBool(m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Item("multiple"))
    Multiple = prop_multiple
End Property
Public Property Get Styles() As Collection
    Set prop_styles = m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Item("styles")
    Set Styles = prop_styles
End Property
Public Property Get OptionalHeadingStyles() As Collection
    If m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Exists("optional_heading_styles") Then
        Set prop_optional_heading_styles = m_dictValues("contiguous_block_criteria_" & _
        m_str_criteria_count).Item("optional_heading_styles")
    Else
        Set prop_optional_heading_styles = New Collection
    End If
    Set OptionalHeadingStyles = prop_optional_heading_styles
End Property
Public Property Get FirstChild() As Boolean
    If m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Exists("first_child") Then
      prop_first_child = True
    Else
      prop_first_child = False
    End If
    FirstChild = prop_first_child
End Property
Public Property Get FirstChildText() As Collection
    If m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Exists("first_child") Then
      Set prop_first_child_text = m_dictValues("contiguous_block_criteria_" & _
      m_str_criteria_count).Item("first_child").Item("text")
    Else
      Set prop_first_child_text = New Collection
    End If
    Set FirstChildText = prop_first_child_text
End Property
Public Property Get FirstChildMatch() As Boolean
    If m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Exists("first_child") Then
      prop_first_child_match = CBool(m_dictValues("contiguous_block_criteria_" & _
      m_str_criteria_count).Item("first_child").Item("match"))
    Else
      prop_first_child_match = True
    End If
    FirstChildMatch = prop_first_child_match
End Property
Public Property Get RequiredStyles() As Collection
    Set prop_required_styles = m_dictValues("contiguous_block_criteria_" & _
    m_str_criteria_count).Item("previous_sibling").Item("required_styles")
    Set RequiredStyles = prop_required_styles
End Property
Public Property Get PreviousUntil() As Collection
    If m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Exists("previous_until") Then
        Set prop_previous_until = m_dictValues("contiguous_block_criteria_" & _
        m_str_criteria_count).Item("previous_until")
    Else
        Set prop_previous_until = New Collection
    End If
    Set PreviousUntil = prop_previous_until
End Property
Public Property Get LastCriteria() As Boolean
    ' "prop_last" value is to let us know if there are more rules coming for this SS;
    ' this is important to know when 'section_required' = true
    If m_dictValues.Exists("contiguous_block_criteria_" & m_str_next_criteria_count) Then
        prop_last = False
    Else
        prop_last = True
    End If
    LastCriteria = prop_last
End Property
Public Property Get Priority() As Long
    If SectionRequired = True Then
        prop_priority = 1
    ElseIf m_dictValues.Exists("order") And m_dictValues.Item("order") = "last" Then
        prop_priority = 4
    ElseIf Position <> vbNullString Then
        prop_priority = 3
    Else
        prop_priority = 2
    End If
    Priority = prop_priority
End Property

' *****************************************************************************
'       PUBLIC METHODS
' *****************************************************************************

' ===== Init ========================================================
' To create new objects for this class, from calls from the factory
' We uses the 'rulenum' to check against contiguous criteria blocks and get value for 'LastCriteria'

Public Sub Init(Name As String, values As Dictionary, rulenum As Long)

    m_str_criteria_count = GetTwoDigitString(rulenum)
    m_str_next_criteria_count = GetTwoDigitString(rulenum + 1)
    Set m_dictValues = values
    prop_ss_name = Name
    
End Sub

' *****************************************************************************
'       PRIVATE METHODS
' *****************************************************************************

' ===== GetTwoDigitString ========================================================
' This function just converts a single digit to a string with a leading "0"

Private Function GetTwoDigitString(num As Long)

    If num < 9 Then
        GetTwoDigitString = "0" & num
    Else
        GetTwoDigitString = CStr(num)
    End If
End Function

