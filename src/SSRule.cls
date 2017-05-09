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
Private prop_SectionStartParas As Collection
Private prop_Blocks As Collection


' module/class only var declarations
Private m_dictValues As Dictionary
Private m_name As String
Private m_str_criteria_count As String
Private m_str_next_criteria_count As String

' *****************************************************************************
'       PUBLIC PROPERTIES
' *****************************************************************************

' *****************************************************************************
'       PRIVATE PROPERTIES
' *****************************************************************************
' all of these are Read-Only, and refer to their equivalent values in the SSRules JSON
' Private because .Process method handles all of these within the class.

Private Property Get RuleName() As String
    prop_rule_name = prop_ss_name & "_" & m_str_criteria_count
    RuleName = prop_rule_name
End Property


Private Property Get SectionName() As String
    SectionName = prop_ss_name
End Property


Private Property Get SectionRequired() As Boolean
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


Private Property Get InsertBefore() As Collection
    If m_dictValues.Exists("section_required") Then
      Set prop_insert_before = m_dictValues("section_required").Item("insert_before")
    Else
      Set prop_insert_before = New Collection
    End If
    InsertBefore = prop_insert_before
End Property


Private Property Get Position() As String
    If m_dictValues.Exists("position") Then
      prop_position = m_dictValues("position")
    Else
      prop_position = ""
    End If
    Position = prop_position
End Property


Private Property Get Multiple() As Boolean
    prop_multiple = CBool(m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Item("multiple"))
    Multiple = prop_multiple
End Property
Public Property Get Styles() As Collection
    Set prop_styles = m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Item("styles")
    Set Styles = prop_styles
End Property


Private Property Get OptionalHeadingStyles() As Collection
    If m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Exists("optional_heading_styles") Then
        Set prop_optional_heading_styles = m_dictValues("contiguous_block_criteria_" & _
        m_str_criteria_count).Item("optional_heading_styles")
    Else
        Set prop_optional_heading_styles = New Collection
    End If
    Set OptionalHeadingStyles = prop_optional_heading_styles
End Property
 
Private Property Get FirstChild() As Boolean
    If m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Exists("first_child") Then
      prop_first_child = True
    Else
      prop_first_child = False
    End If
    FirstChild = prop_first_child
End Property


Private Property Get FirstChildText() As Collection
    If m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Exists("first_child") Then
      Set prop_first_child_text = m_dictValues("contiguous_block_criteria_" & _
      m_str_criteria_count).Item("first_child").Item("text")
    Else
      Set prop_first_child_text = New Collection
    End If
    Set FirstChildText = prop_first_child_text
End Property


Private Property Get FirstChildMatch() As Boolean
    If m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Exists("first_child") Then
      prop_first_child_match = CBool(m_dictValues("contiguous_block_criteria_" & _
      m_str_criteria_count).Item("first_child").Item("match"))
    Else
      prop_first_child_match = True
    End If
    FirstChildMatch = prop_first_child_match
End Property


Private Property Get RequiredStyles() As Collection
    Set prop_required_styles = m_dictValues("contiguous_block_criteria_" & _
    m_str_criteria_count).Item("previous_sibling").Item("required_styles")
    Set RequiredStyles = prop_required_styles
End Property


Private Property Get PreviousUntil() As Collection
    If m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Exists("previous_until") Then
        Set prop_previous_until = m_dictValues("contiguous_block_criteria_" & _
        m_str_criteria_count).Item("previous_until")
    Else
        Set prop_previous_until = New Collection
    End If
    Set PreviousUntil = prop_previous_until
End Property


Private Property Get LastCriteria() As Boolean
    ' "prop_last" value is to let us know if there are more rules coming for this SS;
    ' this is important to know when 'section_required' = true
    If m_dictValues.Exists("contiguous_block_criteria_" & m_str_next_criteria_count) Then
        prop_last = False
    Else
        prop_last = True
    End If
    LastCriteria = prop_last
End Property


Private Property Get Priority() As Long
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



' ===== SectionStartParas =====================================================
' Collection of paragraph indices of paragraphs matching .SectionName property
' as current style. If not un use, returns empty collection, which you can test
' with Collection.Count = 0.

Private Property Get SectionStartParas() As Collection
' If property hasn't been created yet, create it first.
  If prop_SectionStartParas Is Nothing Then
    Dim col_Sections As Collection
    Set col_Sections = New Collection
    Set prop_SectionStartParas = FindStyleInDoc(SearchStyle:=SectionName, _
      DestinationCollection:=col_Sections)
  End If
  
  Set SectionStartParas = prop_SectionStartParas
End Property


' ===== Blocks ================================================================
' Collection of paragraph indices of the starting paragraph of each block of
' contiguous paragraphs that we're editing.

Private Property Get Blocks() As Collection
' Create collection if it doesn't exist yet
  If prop_Blocks Is Nothing Then
    Set prop_Blocks = New Collection
  ' Create new collection, easier to add unknown number of items to
    Dim col_PossibleBlocks As Collection
    Set col_PossibleBlocks = New Collection

  ' Create collection of ALL paragraphs using styles in Styles property
    Dim varStyle As Variant
    For Each varStyle In Styles
      FindStyleInDoc SearchStyle:=varStyle, DestinationCollection:=colPossibleBlocks
    Next varStyle

  ' Sort indices in ascending order. No native Collection.Sort method, convert
  ' to array and sort that inatead
    Dim arr_PossibleBlocks As Variant
    arr_PossibleBlocks = Utils.ToArray(col_PossibleBlocks)
    WordBasic.SortArray (arr_PossibleBlocks)
    Set col_PossibleBlocks = Utils.ToCollection(arr_PossibleBlocks)

  ' Pull out contiguous indices
    Dim lngCurrentIndex As Long
    Dim lngNextIndex As Long
    Dim lngCurrentValue As Long
    Dim lngNextValue As Long
    
    lngCurrentIndex = 1
    
    Do While lngCurrentIndex <= col_PossibleBlocks.Count
      lngCurrentValue = col_PossibleBlocks.Item(lngCurrentIndex)
      prop_Blocks.Add lngCurrentValue
      lngNextIndex = lngCurrentIndex + 1
      
      Do While lngNextIndex <= col_PossibleBlocks.Count
        lngNextValue = col_PossibleBlocks.Item(longnextindex)
        If lngNextValue = lngCurrentValue + 1 Or Utils.ExistsInCollection(VersatileBlockCollection, lngNextValue) = True Then
          lngCurrentValue = lngNextValue
          lngNextIndex = lngNextIndex + 1
        Else
          Exit Do
        End If
      Loop
      lngCurrentIndex = lngNextIndex
    Loop
  End If
  
  Set Blocks = prop_Blocks
End Property

' *****************************************************************************
'       PUBLIC METHODS
' *****************************************************************************

' ===== Init ========================================================
' To create new objects for this class, from calls from the factory
' We uses the 'rulenum' to check against contiguous criteria blocks and get value for 'LastCriteria'

Public Sub Init(Name As String, Values As Dictionary, rulenum As Long)

    m_str_criteria_count = GetTwoDigitString(rulenum)
    m_str_next_criteria_count = GetTwoDigitString(rulenum + 1)
    Set m_dictValues = Values
    prop_ss_name = Name
    
End Sub


' ===== Process ===============================================================
' Check this rule's criteria against activeDoc to see if any section-start paras
' might be missing.

Public Sub Process()
  If SectionStartParas.Count > 0 Then
    
  Else
  
  End If
End Sub

' *****************************************************************************
'       PRIVATE METHODS
' *****************************************************************************

' ===== FindStyleInDoc =========================================================
' Adds index numbers for paragraphs with SearchStyle applied to DestinationCollection.
' Note that because DestinationCollection is passed ByRef we don't need to return
' it to the calling function.

Private Sub FindStyleInDoc(SearchStyle As String, ByRef DestinationCollection _
  As Collection)
  Dim lngFoundIndex As Long

' MacroHelpers.ParaIndex works with current Selection, so use Selection.Find
' Start from beginning of document
  activeDoc.Selection.HomeKey Unit:=wdStory

  MacroHelpers.zz_clearFind
  With Selection.Find
    .Format = True
    .Style = SearchStyle
    .Forward = True
    .Wrap = wdFindStop
    .Execute
    
    Do While .Execute = True
      .Execute
      lngFoundIndex = MacroHelpers.ParaIndex(UseEnd:=False)
      DestinationCollection.Add lngFoundIndex
    Loop
  End With
End Sub


' ===== CountSections ============================================================
' Validate currently used instances of this SectionName and/or Styles.
'


' ===== GetTwoDigitString ========================================================
' This function just converts a single digit to a string with a leading "0"

Private Function GetTwoDigitString(num As Long)

    If num < 9 Then
        GetTwoDigitString = "0" & num
    Else
        GetTwoDigitString = CStr(num)
    End If
End Function

