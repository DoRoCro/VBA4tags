Attribute VB_Name = "TestFailureCodes"
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
Public Sub TestFailureCodeProperties()
    On Error GoTo TestFail
    
    'Arrange:
    Dim fcode As clsFailureCode
    Set fcode = New clsFailureCode
    'Act:
    fcode.Description = "Test Description"
    fcode.ID = "FA_TEST"
    Set fcode.OutputSheet = ThisWorkbook.Worksheets("Output")
    'Assert:
    Assert.Istrue (fcode.ID = "FA_TEST")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub

'@TestMethod
Public Sub TestFailureCodesAdd()
    On Error GoTo TestFail
    
    'Arrange:
    Dim fcodes As clsFailureCodes
    Set fcodes = New clsFailureCodes
    Dim fcode As clsFailureCode
    Set fcode = New clsFailureCode
    
    'Act:
    fcode.Description = "Test Description"
    fcode.ID = "FA_TEST"
    Set fcode.OutputSheet = ThisWorkbook.Worksheets("Output")
    'Assert:
    Assert.Istrue (fcode.ID = "FA_TEST")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub

'@TestMethod
Public Sub TestFailureCodesLoad()
    On Error GoTo TestFail
    
    'Arrange:
    Dim fcodes As clsFailureCodes
    Set fcodes = New clsFailureCodes
    Dim fcode As clsFailureCode
    Set fcode = New clsFailureCode
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("TestDefaultCriticalities").ListObjects("TestFailureCodeDefaultCriticalitiesTable")
    'Act:
    fcodes.LoadTable tbl
    'Assert:
    Assert.Istrue (fcodes.Item(2).ID = "FA_CFSGA")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub


