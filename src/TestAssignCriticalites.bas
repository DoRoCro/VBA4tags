Attribute VB_Name = "TestAssignCriticalites"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
'Public CriticalityWbName As String


'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    CriticalityWbName = ThisWorkbook.Name
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
Public Sub TestCalculateCriticalityGetsRightWorksheet()
    On Error GoTo TestFail
    'Arrange
    Dim tag As clsTag
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim x As Long          ' >32000 entries
    
    Set tag = New clsTag
    With tag
        .Description = "Test Tag"
        .ID = "XYZ-1234"
        .FailureCode = "TestFailureCodeTemplate"
        .Discipline = "INST"
        .SystemID = "78"
    End With
    
    'Act:
    'SetTagCriticalityByFailureCode tag, New clsMAHlist
    tag.SetDefaultCriticalityByFailureCode
    'Assert
    'Assert.istrue (ws.Name = "FA_CFBC")
    'Assert.inconclusive
    Assert.Istrue ("A" = tag.Criticality)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub

