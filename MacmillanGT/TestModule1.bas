Attribute VB_Name = "TestModule1"
Option Explicit

Option Private Module

' Macmillan customization:
Dim docTestfileGood As Document
Dim docTestfileBad As Document

'@TestModule
Private Assert As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")

End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod
Public Sub Test_CheckFileName()
    On Error GoTo TestFail
    
    'Arrange:
        ' Load test docs - they are Dim's as declarations for this module.
        Set docTestfileGood = LoadTestDoc("good")  'param = 'good' or 'bad^', which test doc to load
        Set docTestfileBad = LoadTestDoc("bad^")  'param = 'good' or 'bad^', which test doc to load

        ' Declarations for this test:
        Dim boolTestGood As Boolean
        Dim boolTestBad As Boolean
    
    'Act:
        ' this construtction is necessary to run Private subs/functions from other modules:
        '   Application.Run("Module.Sub")
        ' Looks like if you're passing parameters its weird too:
        '   Application.Run "Test", Variable1, Variable2
        ' I had to add the parentheses to get a function returning a value to work, so will need to test
        ' re: functions with parameters
        docTestfileGood.Activate
        boolTestGood = Application.Run("Reports.CheckFileName")
        docTestfileBad.Activate
        boolTestBad = Application.Run("Reports.CheckFileName")
        
        ' Cleanup - delete test docs
        Call DeleteTestDoc(docTestfileGood.FullName)
        Call DeleteTestDoc(docTestfileBad.FullName)
    
    'Assert:
        Assert.IsFalse boolTestGood
        ' assert.IsTrue would have worked here too:
        Assert.AreEqual boolTestBad, True
        

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



