VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SSRuleCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *****************************************************************************
'       PRIVATE VARIABLES
' *****************************************************************************

Private prop_RuleDictionary As Dictionary
Private prop_sectionLists As Dictionary
Private prop_FrontmatterStyles As Collection
Private prop_MainStyles As Collection
Private prop_BackmatterStyles As Collection
Private prop_rules As Collection
Private prop_Warnings As Collection

' *****************************************************************************
'       PRIVATE PROPERTIES
' *****************************************************************************

' ===== SectionLists[PropertyGet] ========================================================
' This Property creates a dict containing 4 Collections as a Read-Only property, featuring:
' A list(Collection) of "all" Section Start Stylenames, a list of "frontmatter" ones, "backmatter" ones
' and "main" ones.  The "all" Collection is needed when checking for acceptable "previous_siblings"
' The other three are needed when testing "position" criteria

' Set to Private b/c only needed by Validate procedure.

' TODO: Get some benchmarks on how long this takes to create. If more than a couple of seconds,
' move this data to vba_style_config.json.

Private Property Get SectionLists() As Dictionary
' Check if saved variable already exists, else it will go through the code to create
' it every time it is accessed

  If prop_sectionLists Is Nothing Then
    Set prop_sectionLists = New Dictionary
    Dim collFrontmatter As Collection
    Dim collMain As Collection
    Dim collBackmatter As Collection
    Dim collAll As Collection
    Dim J As Long
    
    ' define the collections & dict
    Set collFrontmatter = New Collection
    Set collMain = New Collection
    Set collBackmatter = New Collection
    Set collAll = New Collection
    
    ' cycle through the json data and add section start names to appropriate collections
    Dim strSSname As String
    For J = 0 To RuleDictionary.Count - 1
        strSSname = RuleDictionary.Keys(J)
        Debug.Print "style name: " & strSSname
        collAll.Add (strSSname)
        If RuleDictionary(strSSname).Item("section_type") = "frontmatter" Then
            collFrontmatter.Add (strSSname)
        ElseIf RuleDictionary(strSSname).Item("section_type") = "main" Then
            collMain.Add (strSSname)
        ElseIf RuleDictionary(strSSname).Item("section_type") = "backmatter" Then
            collBackmatter.Add (strSSname)
        End If
    Next
    
    ' add collections to the dict
    prop_sectionLists.Add "frontmatter", collFrontmatter
    prop_sectionLists.Add "main", collMain
    prop_sectionLists.Add "backmatter", collBackmatter
    prop_sectionLists.Add "all", collAll
  End If
    
  Set SectionLists = prop_sectionLists
    
End Property


' ===== Rules[PropertyGet] ========================================================
' This Property creates a Collection of SSRule objects, by cycling through the json data
' and passing contiguous-block-criteria for each SectionStart (there may be more than one)
' to the SSRule factory, and adding the returned SSRUle to the prop_rules collection

Private Property Get Rules() As Collection
' Check if it's already been created first
  If prop_rules Is Nothing Then
    Set prop_rules = New Collection
    Dim dictSingleSectionStart As Dictionary
    Dim objNewSectionStartRule As SSRule
    Dim lng_criteriaCount As Long
    Dim strCriteriaKey As String
    Dim varKey As Variant
    Dim lngPriority As Long

  ' Set up priority collections
  ' Could probably be more flexible (in case more than 4 priorities) but works for now
  ' Index will match .Priority
    Dim A As Long
    Dim colPriorty As Collection
    For A = 1 To 4
      Set colPriorty = New Collection
      prop_rules.Add colPriorty
    Next A
    
    ' Cycle through the JSON data
    Dim strSSname As String
    For Each varKey In RuleDictionary.Keys
        strSSname = varKey
        ' get / set the value of this Section Start item as a dict
        Set dictSingleSectionStart = RuleDictionary(varKey)
    ' Conditional so we don't create a Rule object for an entry with no criteria
'        If dictSingleSectionStart.Exists("contiguous_block_criteria_01") Then
'            ' Create the rule object & add it to rule object collection
'            Set objNewSectionStartRule = Factory.CreateSSrule(varKey, dictSingleSectionStart, 1) ', prop_sectionLists)
'            prop_rules.Add objNewSectionStartRule
        
        ' Can use this for whole loop instead of block above, b/c checking .Exists before creating Rule
        lng_criteriaCount = 1
        strCriteriaKey = "contiguous_block_criteria_" & Format(lng_criteriaCount, "00")
        Do While dictSingleSectionStart.Exists(strCriteriaKey)
            ' Create the rule object
            Set objNewSectionStartRule = Factory.CreateSSrule(strSSname, dictSingleSectionStart, lng_criteriaCount)
            lngPriority = objNewSectionStartRule.Priority
            ' add its name + priority to the prop_rules collection
            prop_rules(lngPriority).Add objNewSectionStartRule
            ' increment count & string
            lng_criteriaCount = lng_criteriaCount + 1
            strCriteriaKey = "contiguous_block_criteria_" & Format(lng_criteriaCount, "00")
        Loop
'        End If
    Next varKey
  End If
    
  Set Rules = prop_rules
    
End Property


' ===== RuleDictionary ========================================================
' Separate property for the Dictionary containing the JSON data, including
' downloading the file. Also better encapsulates the creation of the dict, in
' case we change the source in the future.

' PARAMS:
' JsonFileName : Just file name w/ extension. Rest of info (where to download
' from, where to save locally, etc.) is created in this procedure.

Private Property Get RuleDictionary() As Dictionary
' Check if already created before we do it again.
  If prop_RuleDictionary Is Nothing Then
    Dim strJsonPath As String
  
  ' Download JSON file
    strJsonPath = SharedFileInstaller.DownloadJson("section_start_rules.json")
  
  ' Load JSON data into dictionary
    Set prop_RuleDictionary = ClassHelpers.ReadJson(strJsonPath)
  End If
  
  Set RuleDictionary = prop_RuleDictionary
End Property


' ===== Warnings ==============================================================
' Collection of dictionaries with data about missing section starts. Use
' native Collection.Add method to add, use WarningString property to return
' actual text for Style Report.

Private Property Get Warnings() As Collection
  If prop_Warnings Is Nothing Then
    Set prop_Warnings = New Collection
  End If

  Set Warnings = prop_Warnings
End Property

' ===== WarningString =========================================================
' Converts Warnings collection to a single string to print in Report output.

Private Property Get WarningString() As String
  Dim strOutput As String
  If Warnings.Count > 0 Then
    strOutput = _
    "** WARNING: Please check if you're missing a Section style on these pages:" & vbNewLine

    Dim varWarning As Variant
    For Each varWarning In Warnings
      strOutput = strOutput & vbNewLine & _
      "            pg. " & varWarning("page_number") & ": " & varWarning("missing_style")
    Next varWarning
    
    strOutput = strOutput & vbNewLine & _
    "HINT: Each book section needs to start with an extra paragraph that has a " & vbNewLine & _
    Chr(34) & "Section-Start" & Chr(34) & " style applied to it. If the pages listed above start new sections" & vbNewLine & _
    "but don't have a Section style, you should add one now."
  Else
    strOutput = vbNullString
  End If
  WarningString = strOutput
End Property


' *****************************************************************************
'       PUBLIC METHODS
' *****************************************************************************

' ===== Validate ==============================================================
' Loops through .Rules to run validation checks. Returns single string Error
' message; if no errors, returns null string.

Public Function Validate() As String
  Dim objRule As Object
  Dim specificRule As SSRule
  Dim A As Long
  For A = 1 To 4
    For Each objRule In Rules(A)
      Set specificRule = objRule
      With specificRule
      ' Check .Position property separately so we can pass SectionLists collection
      ' from here.
        If .Position <> WT_SectionType.no_section Then
          .CheckPosition Sections:=SectionLists
        End If
        .Process
        If .Blocks.Count > 0 Then ' we have possible missing ss-starts
          AddWarning ProcessedRule:=specificRule
        End If
      End With
    Next objRule
  Next A
  
  Dim strFinalWarning As String
  strFinalWarning = WarningString

  Validate = strFinalWarning
End Function


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

' ===== AddWarning ============================================================
' Create dictionary of missing ss-style warning info and add to Warnings property.

Private Sub AddWarning(ProcessedRule As SSRule)
  Dim varBlock As Variant
  Dim dictWarningData As Dictionary
  Dim lngPageNumber As Long
  Dim lngIndex As Long

  With ProcessedRule
    For Each varBlock In .Blocks
      lngIndex = varBlock
      Set dictWarningData = New Dictionary
    ' Get page number of first para of block
      lngPageNumber = MacroHelpers.ParaInfo(paraInd:=lngIndex, _
        InfoType:=wdActiveEndAdjustedPageNumber)
  
      dictWarningData.Add "page_number", lngPageNumber
      dictWarningData.Add "missing_style", .SectionName
  
      Warnings.Add dictWarningData
    Next varBlock
  End With
End Sub
