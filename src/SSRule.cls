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
Private prop_position As WT_SectionType
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
Private prop_SectionStartsInUse As Collection
Private prop_Blocks As Collection
Private prop_MissingWarnings As Collection


' module/class only var declarations
Private m_dictValues As Dictionary
Private m_name As String
Private m_str_criteria_count As String
Private m_str_next_criteria_count As String

' *****************************************************************************
'       PUBLIC PROPERTIES
' *****************************************************************************

Public Property Get SectionName() As String
    SectionName = prop_ss_name
End Property


Public Property Get Priority() As Long
  If prop_priority = 0 Then
    If SectionRequired = True Then
        prop_priority = 1
    ElseIf m_dictValues.Exists("order") And m_dictValues.Item("order") = "last" Then
        prop_priority = 4
    ElseIf Position <> WT_SectionType.no_section Then
        prop_priority = 3
    Else
        prop_priority = 2
    End If
  End If
  Priority = prop_priority
End Property


' ===== Blocks ================================================================
' Collection of paragraph indices of the starting paragraph of each block of
' contiguous paragraphs that we're editing.

' Includes a PRIVATE Property Set so we can manually add an InsertBefore index,
' or to set to New or Nothing, as needed.

Public Property Get Blocks() As Collection
' Create collection if it doesn't exist yet
  If prop_Blocks Is Nothing Then
    Set prop_Blocks = New Collection
  ' Create new collection, easier to add unknown number of items to
    Dim col_PossibleBlocks As Collection
    Set col_PossibleBlocks = New Collection

  ' Create collection of ALL paragraphs using styles in Styles property
    Dim varStyle As Variant
    For Each varStyle In Styles
      ClassHelpers.FindStyleIndex CStr(varStyle), col_PossibleBlocks, Multiple
    Next varStyle

  ' Sort indices in ascending order.
    If col_PossibleBlocks.Count > 0 Then
      Utils.SortCollection col_PossibleBlocks
    End If

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
        lngNextValue = col_PossibleBlocks.Item(lngNextIndex)
        If lngNextValue = (lngCurrentValue + 1) Or _
          WT_StyleConfig.VersatileStyleExistsByIndex(lngNextValue) = True Then
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

Public Property Get Position() As WT_SectionType
    If m_dictValues.Exists("position") Then
      prop_position = WT_Enums.CSectionType(m_dictValues("position"))
    Else
      prop_position = WT_SectionType.no_section
    End If
    Position = prop_position
End Property

' *****************************************************************************
'       PRIVATE PROPERTIES
' *****************************************************************************

Private Property Set Blocks(Value_Collection As Collection)
  Set prop_Blocks = Value_Collection
End Property

' all of these are Read-Only, and refer to their equivalent values in the SSRules JSON
' Private because .Process method handles all of these within the class.

Private Property Get RuleName() As String
    prop_rule_name = prop_ss_name & "_" & m_str_criteria_count
    RuleName = prop_rule_name
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
  If prop_insert_before Is Nothing Then
    If m_dictValues.Exists("section_required") Then
      Set prop_insert_before = m_dictValues("section_required").Item("insert_before")
    Else
      Set prop_insert_before = New Collection
    End If
  End If
  InsertBefore = prop_insert_before
End Property




Private Property Get Multiple() As Boolean
    prop_multiple = CBool(m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Item("multiple"))
    Multiple = prop_multiple
End Property

Private Property Get Styles() As Collection
  If prop_styles Is Nothing Then
    Set prop_styles = m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Item("styles")
  End If
  Set Styles = prop_styles
End Property

Private Property Get OptionalHeadingStyles() As Collection
  If prop_optional_heading_styles Is Nothing Then
    If m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Exists("optional_heading_styles") Then
        Set prop_optional_heading_styles = m_dictValues("contiguous_block_criteria_" & _
        m_str_criteria_count).Item("optional_heading_styles")
    Else
        Set prop_optional_heading_styles = New Collection
    End If
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
  If prop_first_child_text Is Nothing Then
    If m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Exists("first_child") Then
      Set prop_first_child_text = m_dictValues("contiguous_block_criteria_" & _
      m_str_criteria_count).Item("first_child").Item("text")
    Else
      Set prop_first_child_text = New Collection
    End If
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
  If prop_required_styles Is Nothing Then
    Set prop_required_styles = m_dictValues("contiguous_block_criteria_" & _
    m_str_criteria_count).Item("previous_sibling").Item("required_styles")
  End If
  Set RequiredStyles = prop_required_styles
End Property

Private Property Get PreviousUntil() As Collection
  If prop_previous_until Is Nothing Then
    If m_dictValues("contiguous_block_criteria_" & m_str_criteria_count).Exists("previous_until") Then
        Set prop_previous_until = m_dictValues("contiguous_block_criteria_" & _
        m_str_criteria_count).Item("previous_until")
    Else
        Set prop_previous_until = New Collection
    End If
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



' ===== SectionStartsInUse =====================================================
' Collection of paragraph indices of paragraphs matching .SectionName property
' as current style. If not un use, returns empty collection, which you can test
' with Collection.Count = 0.

Private Property Get SectionStartsInUse() As Collection
' If property hasn't been created yet, create it first.
  If prop_SectionStartsInUse Is Nothing Then
    Set prop_SectionStartsInUse = New Collection
    ClassHelpers.FindStyleIndex SectionName, prop_SectionStartsInUse, True
  End If
  
  Set SectionStartsInUse = prop_SectionStartsInUse
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

' Checking the sections is the only part that checks once per Rule, not per Block
' CheckSectionCounts = True means to continue validation. Each subsequent method
' removes items from Blocks if they don't need more validation, so once finished
' only errors will remain in Blocks.

Public Sub Process()
  If CheckSectionCounts = True Then
' Moved CheckPosition to calling method in parent, so we don't have to pass
' SectionLists collection to every rule.
'    CheckPosition
    AddOptionalHeadings
    CheckPreviousSiblings
    CheckFirstChild
  End If
End Sub



' ===== CheckPosition =========================================================
' Validate Position property. If validation determines a contiguous block is
' OK as is, it is removed from the Blocks collection.

' Public so it can be run from calling procedure: We can check if this needs
' to be run at all before we call it, to avoid having to pass the Sections
' collection to every rule when this property is only used once.

Public Sub CheckPosition(Sections As Dictionary)
  If Position <> WT_SectionType.no_section Then
    Dim varIndex As Variant
    Dim lngSectionStart As Long
    Dim lngSectionEnd As Long

    Dim col_Temp As Collection
    Set col_Temp = New Collection
  
    Select Case Position
      Case WT_SectionType.frontmatter
        lngSectionStart = FindSectionStart("frontmatter", Sections)
        lngSectionEnd = FindSectionStart("main", Sections)
        If lngSectionEnd = 0 Then
          lngSectionEnd = FindSectionStart("backmatter", Sections)
        End If
        If lngSectionEnd > 0 Then
          lngSectionEnd = lngSectionEnd - 1
        End If
        
      Case WT_SectionType.main
        lngSectionStart = FindSectionStart("main", Sections)
        lngSectionEnd = FindSectionStart("backmatter", Sections)
        If lngSectionEnd > 0 Then
          lngSectionEnd = lngSectionEnd - 1
        End If

      Case WT_SectionType.backmatter
        lngSectionStart = FindSectionStart("backmatter", Sections)
        lngSectionEnd = activeDoc.Paragraphs.Count

    End Select
  
  ' If no section start, clearly we don't have that section
    Dim lngIndex As Long
    If CBool(lngSectionStart) Then
      For Each varIndex In Blocks
        lngIndex = varIndex
      ' Check if block paragraph is after the start paragraph
        If lngIndex >= lngSectionStart Then
        ' If no end, following section doesn't exist
          If CBool(lngSectionEnd) Then
          ' Check if block is before end of section
            If lngIndex <= lngSectionEnd Then
              col_Temp.Add varIndex
            End If
          End If
        End If
      Next varIndex
    End If
    
    Set Blocks = col_Temp
  
  End If
End Sub
' *****************************************************************************
'       PRIVATE METHODS
' *****************************************************************************

' ===== CheckSectionCounts ====================================================
' Only Process step that checks the Rule. Returns True if should keep validating
' and returns False if should stop validating. Also sets Blocks property, from
' which we'll pull errors later.

Private Function CheckSectionCounts() As Boolean
  CheckSectionCounts = False
  If SectionStartsInUse.Count > 0 Then
    If Multiple = False Then
    ' Multiple = False and we found at least 1 section, so stop processing rule
    ' Only condition where Blocks is still Nothing; must set here b/c final
    ' Blocks.Count determines errors, but if we Get Blocks when it is Nothing,
    ' it will run the search!
      Set Blocks = New Collection
    Else ' Multiple = True
      If SectionStartsInUse.Count < Blocks.Count Then
        ' Only continue if more blocks than sections
        CheckSectionCounts = True
      End If
    End If
  Else  ' SectionStartsInUse = 0
    If Blocks.Count = 0 Then
    ' No section-starts *OR* blocks, which is only an issue if SectionRequired
    ' (but either way don't continue, no Blocks to check)
      If SectionRequired = True Then
        Dim col_FirstInd As Collection
        Set col_FirstInd = New Collection
        col_FirstInd.Add CheckInsertBefore()
        Set Blocks = col_FirstInd
      End If
    End If
  End If
End Function

' ===== CheckInsertBefore =====================================================
' If InsertBefore property exists, finds first instance of style in collection
' and returns the paragraph index.

Private Function CheckInsertBefore() As Long
  If InsertBefore.Count > 0 Then
  ' Create collection of found style
    Dim varStyle As Variant
    Dim col_InsertBeforeParas As Collection
    Set col_InsertBeforeParas = New Collection

    For Each varStyle In InsertBefore
    ' ReturnMultiple = False means will only return ONE result for each style
      ClassHelpers.FindStyleIndex varStyle, col_InsertBeforeParas, False
    Next varStyle

    Dim lngFoundPara As Long
    If col_InsertBeforeParas.Count = 0 Then
    ' Enter warning for first paragraph
      col_InsertBeforeParas.Add 1
    ElseIf col_InsertBeforeParas.Count > 1 Then
    ' Sort so lowest paragraph index is first
      Utils.SortCollection col_InsertBeforeParas
    End If

    CheckInsertBefore = col_InsertBeforeParas.Item(1)

  End If
End Function

' ===== FindSectionStart ======================================================
' Finds first paragraph in doc with a style for that section type. Options are
' "frontmatter", "main", and "backmatter".

' PARAMS
' SectionType: type of section we're looking for (fm, main, bm)
' SectionStyles: Collection of section-start styles for that type.

' RETURNS:
' Paragraph index of first instance, or 0 if none found.

Private Function FindSectionStart(SectionType As String, SectionStyles As _
  Dictionary) As Long

' Pull out collection of styles for our section type
  Dim col_SearchStyles As Collection
  Set col_SearchStyles = SectionStyles(SectionType)

  Dim varStyle As Variant
  Dim col_FirstInstance As Collection
  Set col_FirstInstance = New Collection

' Create collection of the FIRST instance of each section style
  For Each varStyle In col_SearchStyles
    ClassHelpers.FindStyleIndex CStr(varStyle), col_FirstInstance, False
  Next varStyle

  If col_FirstInstance.Count > 0 Then
    Utils.SortCollection col_FirstInstance
    FindSectionStart = col_FirstInstance(1)
  Else
    FindSectionStart = 0
  End If
  
End Function



' ===== AddOptionalHeadings ===================================================
' If rule has optional headings, check each block and add if found.

Private Sub AddOptionalHeadings()
  If OptionalHeadingStyles.Count > 0 Then
    Dim varIndex As Variant
    Dim lngCheckPara As Long
    Dim lngBlockStart As Long
    Dim strStyle As String
    Dim col_Temp As Collection
    Set col_Temp = New Collection
    
    For Each varIndex In Blocks
    ' Current starting paragraph for our block
      lngBlockStart = varIndex
    ' Paragraph we're checking
      lngCheckPara = varIndex - 1
      Do While lngCheckPara > 0
        strStyle = ActiveDocument.Paragraphs(lngCheckPara).Range.ParagraphStyle
        If Utils.ExistsInCollection(OptionalHeadingStyles, strStyle) = True Or _
          WT_StyleConfig.IsVersatileStyle(strStyle) = True Then
          ' Previous para IS an optional heading, so adjust our validated Block Start
          lngBlockStart = lngCheckPara
          ' Adjust next check para here, since test is at start of Do loop
          lngCheckPara = lngCheckPara - 1
        Else
        ' If didn't match, we've found our block start
          Exit Do
        End If
      Loop
    ' Add all block starts to new collection (if nothing matched, it will
    ' be the same as before
      col_Temp.Add lngBlockStart
    Next varIndex
  ' Replace Blocks property with our adjusted paragraph indices
    Set Blocks = col_Temp
  End If
End Sub

' ===== CheckPreviousSiblings =================================================
' Check "previous_sibling" and/or "previous_until" criteria. If pass (ie., OK as
' is), removes from Blocks collection.

Private Sub CheckPreviousSiblings()
  If Blocks.Count > 0 Then
    Dim col_Temp As Collection
    Set col_Temp = New Collection
    Dim varIndex As Variant
    Dim lngCheckIndex As Long
    Dim strStyle As String
    
    For Each varIndex In Blocks
      lngCheckIndex = varIndex - 1
      Do While lngCheckIndex > 0
        strStyle = activeDoc.Paragraphs(lngCheckIndex).Range.ParagraphStyle
        If Utils.ExistsInCollection(RequiredStyles, strStyle) = True Or _
          WT_StyleConfig.IsSectionStartStyle(strStyle) = True Then
          ' Previous style is OK, exit Do and on to next block
            Exit Do
        ' If no PreviousUntil, this won't match
        ElseIf Utils.ExistsInCollection(PreviousUntil, _
          strStyle) = True Then
          ' Matched PreviousUntil, missing ss style
            col_Temp.Add varIndex
            Exit Do
        Else
          lngCheckIndex = lngCheckIndex - 1
        End If
      
      ' If no PreviousUntil, only continue searching if it's a VersatileStyle
      ' Otherwise it's an error.
        If PreviousUntil.Count = 0 And _
           WT_StyleConfig.IsVersatileStyle(strStyle) = False Then
        ' Definitely didn't match PreviousStyle, so add to new collection of errors
          col_Temp.Add varIndex
          Exit Do
        End If
      
      Loop
    Next varIndex

    Set Blocks = col_Temp
  End If
End Sub

' ===== CheckFirstChild =======================================================
' Validate "first_child" criteria, removes from Blocks if ok as is.

Private Sub CheckFirstChild()
  If FirstChild = True Then
    Dim varIndex As Variant
    Dim col_Temp As Collection
    Set col_Temp = New Collection
    Dim strFirstText As String
    Dim blnMatch As Boolean
    Dim lngIndex As Long
    
    For Each varIndex In Blocks
      lngIndex = varIndex
      strFirstText = MacroHelpers.GetTextByIndex(lngIndex)
      blnMatch = Utils.ExistsInCollection(FirstChildText, strFirstText)
      If blnMatch = FirstChildMatch Then
        col_Temp.Add varIndex
      End If
    Next varIndex
    Set Blocks = col_Temp
  End If
End Sub

' ===== GetTwoDigitString ========================================================
' This function just converts a single digit to a string with a leading "0"

Private Function GetTwoDigitString(num As Long)

    If num < 9 Then
        GetTwoDigitString = "0" & num
    Else
        GetTwoDigitString = CStr(num)
    End If
End Function

