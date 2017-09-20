Attribute VB_Name = "TestAssignCriticalites"
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
Public Sub TestCalculateCriticalityGetsRightWorksheet()
    On Error GoTo TestFail
    'Arrange
    Dim tag As clsTag
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim x As Long          ' >32000 entries
    Set wb = Application.Workbooks("WND Criticality Template.xlsx")
    Set ws = wb.Worksheets("AssetRegisterDefaultCodeApplied")
    Set tag = New clsTag
    With tag
        .Description = "Test Tag"
        .ID = "XYZ-1234"
        .FailureCode = "FA_CFBC"
        .Discipline = "INST"
        .SystemID = "78"
    End With
    
    'Act:
    SetTagCriticalityByFailureCode tag
    'Assert
    'Assert.istrue (Excel.ActiveSheet.Name = "FA_CFBC")
    'Assert.inconclusive
    Assert.AreEqual "A", tag.Criticality

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

