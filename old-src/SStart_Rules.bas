Attribute VB_Name = "SStart_Rules"
Option Explicit

' ===== CreateSectionStartRules ========================================================
' This Sub passes the json filepath to the SSRuleCollection class' factory,
' That class returns a collection of SSRule objects
' Then this sub loops through the returned RuleCollection, in order of increasing
' "priority" value, passing them to the ProcessRule Sub to be processed.

Sub CreateSectionStartRules()
  Dim strJsonFilepath As String
  Dim strJsonFilename As String
  Dim objNewSSruleCollection As SSRuleCollection
  Dim lngRuleCount As Long
  Dim strRuleName As String
  Dim lngRulePriority As Long
  Dim lngPriorityCount As Long
  Dim lngPriorityCheck As Long

  strJsonFilename = "section_start_rules.json"

'    ' TEST PATH (just make a folder in the .docm's path called 'bookmaker_validator' with json file in it)
'    strJsonFilepath = CreateObject("Scripting.FileSystemObject").Getfile(ThisDocument.FullName).ParentFolder.Path _
'    & Application.PathSeparator & "bookmaker_validator" & Application.PathSeparator & strJsonFilename

  '' PRODUCTION PATH (not yet tested :)
  strJsonFilepath = "S:" & Application.PathSeparator & "resources" & Application.PathSeparator & _
  "bookmaker_scripts" & Application.PathSeparator & "bookmaker_validator" & strJsonFilename

  ' create collection object (which creates a collection of Rule objects)
  Set objNewSSruleCollection = Factory.CreateSSRuleCollection(strJsonFilepath)

  ' Loop through Rules by "priority" values (set in SSRule.cls)
  lngPriorityCount = 1
  lngPriorityCheck = 1
  Do Until lngPriorityCheck = 0
    lngPriorityCheck = 0
      For lngRuleCount = 1 To objNewSSruleCollection.Rules.Count
        If objNewSSruleCollection.Rules(lngRuleCount).Priority = lngPriorityCount Then
          Call ProcessRule(objNewSSruleCollection.Rules(lngRuleCount), objNewSSruleCollection.SectionLists)
          lngPriorityCheck = lngPriorityCheck + 1
        End If
      Next
    lngPriorityCount = lngPriorityCount + 1
  Loop

End Sub

' ===== ProcessRule ========================================================
' This would be where the rules would be processed.
' For now I just have debug output here

Sub ProcessRule(p_rule As SSRule, p_sectionLists As Dictionary)

  ' Make sure info from SSRule looks right
  ' Call CheckCollection(p_sectionLists("all"))
  Debug.Print p_rule.Priority & " " & p_rule.RuleName
  'Debug.Print p_rule.SectionName
  'Debug.Print p_rule.SectionRequired
  'Debug.Print p_rule.Position
  'Debug.Print p_rule.Multiple
  'Call CheckCollection(p_rule.Styles)
  'Call CheckCollection(p_rule.OptionalHeadingStyles)
  'Debug.Print p_rule.FirstChild
  'Call CheckCollection(p_rule.FirstChildText)
  'Debug.Print p_rule.FirstChildMatch
  'Call CheckCollection(p_rule.RequiredStyles)
  'Call CheckCollection(p_rule.PreviousUntil)
  'Debug.Print p_rule.LastCriteria

End Sub

' ===== CheckCollection ========================================================
' assisting in output test in ProcessRule!

Private Sub CheckCollection(C As Collection)
    Dim c_item As Variant

    If C.Count > 0 Then
    For Each c_item In C
    Debug.Print "   " & c_item
    Next
    Debug.Print "total item count: " & C.Count
    Else
    Debug.Print "empty collection"
    End If
    
End Sub

