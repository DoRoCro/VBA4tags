Attribute VB_Name = "TestSetupCriticalityTemplate"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
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
Public Sub TestCopyDefaultCriticalitiesIntoTemplateWorksheet()
    On Error GoTo TestFail
    
    'Arrange:
    Dim codeRow As ListRow
    Set codeRow = ThisWorkbook.Worksheets("TestDefaultCriticalities").ListObjects("TestFailureCodeDefaultCriticalitiesTable").ListRows(1)
    'Act:
    Call CopyDefaultCriticalitiesIntoTemplateWorksheet(codeRow, ThisWorkbook.Worksheets("TestFailureCodeTemplate"))
    'Assert:
    Debug.Print ThisWorkbook.Worksheets("TestDefaultCriticalities").Range("B1").Formula
    Assert.IsTrue ThisWorkbook.Worksheets("TestFailureCodeTemplate").Range("B1").Formula = "FA_CFBC"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


