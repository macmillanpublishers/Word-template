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

Private prop_parsedJSON As Dictionary
Private prop_sectionLists As Dictionary
Private prop_rules As Collection
Private strSSname As String

' *****************************************************************************
'       PUBLIC PROPERTIES
' *****************************************************************************

' ===== SectionLists[PropertyGet] ========================================================
' This Property creates a dict containing 4 Collections as a Read-Only property, featuring:
' A list(Collection) of "all" Section Start Stylenames, a list of "frontmatter" ones, "backmatter" ones
' and "main" ones.  The "all" Collection is needed when checking for acceptable "previous_siblings"
' The other three are needed when testing "position" criteria

Public Property Get SectionLists() As Dictionary
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
    Set prop_sectionLists = New Dictionary
    
    ' cycle through the json data and add section start names to appropriate collections
    For J = 0 To prop_parsedJSON.Count - 1
        strSSname = prop_parsedJSON.Keys(J)
        collAll.Add (strSSname)
        If prop_parsedJSON(strSSname).Item("section_type") = "frontmatter" Then
            collFrontmatter.Add (strSSname)
        ElseIf prop_parsedJSON(strSSname).Item("section_type") = "main" Then
            collMain.Add (strSSname)
        ElseIf prop_parsedJSON(strSSname).Item("section_type") = "backmatter" Then
            collBackmatter.Add (strSSname)
        End If
    Next
    
    ' add collections to the dict
    prop_sectionLists.Add "frontmatter", collFrontmatter
    prop_sectionLists.Add "main", collMain
    prop_sectionLists.Add "backmatter", collBackmatter
    prop_sectionLists.Add "all", collAll
    
    Set SectionLists = prop_sectionLists
    
End Property

' ===== Rules[PropertyGet] ========================================================
' This Property creates a Collection of SSRule objects, by cycling through the json data
' and passing contiguous-block-criteria for each SectionStart (there may be more than one)
' to the SSRule factory, and adding the returned SSRUle to the prop_rules collection

Public Property Get Rules() As Collection
    Dim dictSingleSectionStart As Dictionary
    Dim objNewSectionStartRule As SSRule
    Dim i As Long
    Dim lng_criteriaCount As Long
    Dim str_criteriaCount As String
    Dim lngRulePriority As Long
    Dim dictPriority As Dictionary
    
    Set prop_rules = New Collection
    
    ' Cycle through the JSON data
    For i = 0 To prop_parsedJSON.Count - 1
        Set dictSingleSectionStart = New Dictionary
        Set dictPriority = New Dictionary
        ' get / set the name of this Section Start item
        strSSname = prop_parsedJSON.Keys(i)
        ' get / set the value of this Section Start item as a dict
        Set dictSingleSectionStart = prop_parsedJSON(strSSname)
        
        ' Conditional so we don't create a Rule object for an entry with no criteria
        If dictSingleSectionStart.Exists("contiguous_block_criteria_01") Then
            ' Create the rule object & add it to rule object collection
            Set objNewSectionStartRule = Factory.CreateSSrule(strSSname, dictSingleSectionStart, 1) ', prop_sectionLists)
            prop_rules.Add objNewSectionStartRule
            
            ' Add new rules for any extra sets of criteria; start checking at "2"
            lng_criteriaCount = 2
            str_criteriaCount = GetTwoDigitString(lng_criteriaCount)
            Do While prop_parsedJSON(strSSname).Exists("contiguous_block_criteria_" & str_criteriaCount)
                ' Create the rule object
                Set objNewSectionStartRule = Factory.CreateSSrule(strSSname, dictSingleSectionStart, lng_criteriaCount)
                ' add its name + priority to the prop_rules collection
                prop_rules.Add objNewSectionStartRule '.RuleName ', lngRulePriority
                ' increment count & string
                lng_criteriaCount = lng_criteriaCount + 1
                str_criteriaCount = GetTwoDigitString(lng_criteriaCount)
            Loop
        End If
    Next
    
    Set Rules = prop_rules
    
End Property

' *****************************************************************************
'       PUBLIC METHODS
' *****************************************************************************

' ===== Init ========================================================
' This is called by the factory for this class to initialize a new object,
' in this case just reading in the json file from the passed filepath as a dict

Public Sub Init(p_JSONpath As String)

    Set prop_parsedJSON = ClassHelpers.ReadJson(p_JSONpath)

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

    
