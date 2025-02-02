Attribute VB_Name = "TestFailureCodeTemplateLogic"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object
Private wb As Workbook
Private ws As Worksheet


'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TestFailurecodeTemplate")
    ws.Range("B1").Formula = "FA_TESTING"
    ws.Range("B2").Formula = "Unit testing Logic..."
    CriticalityWbName = ThisWorkbook.Name
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
     'Reset all values in template to H3 before starting
    Dim rowStr As Variant
    For Each rowStr In Array("16", "22", "28", "34")
        ws.Range("B" & rowStr).Value = "H"
        ws.Range("C" & rowStr).Value = "3"
    Next
    
    'Set redundancy to no
    ws.Range("G9").Formula = "No"
    
    'Set MAH barriers to blank
    ws.Range("H17").Formula = ""
    ws.Range("I17").Formula = ""
    
    'Set IPL in LOPA flag to No
    ws.Range("J35").Formula = "No"
    'Set Non-Fin regulatory override to No
    ws.Range("I37").Formula = "No"
    
    'set likelihood threshold for business impacts to 5
    ws.Range("J26").Formula = "5"
    
    ws.Range("B1").Formula = "FA_code here"
    ws.Range("B2").Formula = "FA_Code Description here"
    Set ws = Nothing
    Set wb = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
    
    'Reset all values in template to H3 before starting
    Dim rowStr As Variant
    For Each rowStr In Array("16", "22", "28", "34")
        ws.Range("B" & rowStr).Value = "H"
        ws.Range("C" & rowStr).Value = "3"
    Next
    
    'Set redundancy to no
    ws.Range("G9").Formula = "No"
    
    'Set MAH barriers to blank
    ws.Range("H17").Formula = ""
    ws.Range("I17").Formula = ""
    
    'Set IPL in LOPA flag to No
    ws.Range("J35").Formula = "No"
    'Set Non-Fin regulatory override to No
    ws.Range("I37").Formula = "No"
    
    'set likelihood threshold for business impacts to 5
    ws.Range("J26").Formula = "5"
    
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod
Public Sub TestInitialStateGivesC()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:

    'Assert:
    Assert.Istrue (ws.Range("K1") = "C")
    Assert.Istrue (ws.Range("G6").Value = "NO") ' not SCE

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub

'@TestMethod
Public Sub TestSafetyE7GivesAAndNotSCE()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    ws.Range("B16").Formula = "E"
    ws.Range("C16").Formula = "7"
    'Assert:
    Assert.Istrue (ws.Range("K1") = "A")
    Assert.Istrue (ws.Range("G6").Value = "NO")   ' is SCE
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub

'@TestMethod
Public Sub TestSafetyD5GivesAAndSCE()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    ws.Range("B16").Formula = "D"
    ws.Range("C16").Formula = "5"
    
    'Assert:
    Assert.Istrue (ws.Range("K1") = "A")        ' no rule explicitly sets high criticality from safety impact, but expect MAH override in practice
    Assert.Istrue (ws.Range("G6").Value = "YES")   ' is SCE
    Assert.Istrue (ws.Range("G8").Value = 10)   ' CMMS Location priority
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub


'@TestMethod
Public Sub TestSetupForProcess()
    On Error GoTo TestFail
    
    'Arrange:
    Dim fcode As clsFailureCode
    Dim fcodes As clsFailureCodes
    Set fcode = New clsFailureCode
    Set fcodes = New clsFailureCodes
    Dim MAH As clsMAHDefault
    Dim MAHlist As clsMAHlist
    Set MAH = New clsMAHDefault
    Set MAHlist = New clsMAHlist
    
    fcode.Description = "Test Code"
    fcode.ID = "TestFailureCodeTemplate"
    fcodes.Add fcode
    MAH.ID = "TestFailureCodeTemplate"
    MAH.Comment = "Test Comment"
    MAH.Family = "Test Family"
    MAH.Component = "Test Component gives #N/A lookup result so should give #N/A in K17"
    MAHlist.Add MAH
    
    'Act:
    fcodes.SetupForSystemGroup MAHlist
    
    'Assert:
    Assert.Istrue (ws.Range("H17").Text = "Test Family")
    Assert.Istrue (ws.Range("K17").Text = "#N/A")
    Assert.Istrue (ws.Range("I19").Text = "Test Comment")
    Assert.Istrue (ws.Range("K1").Text = "C")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub
'@TestMethod
Public Sub TestCriticalityByFailureCode()
    On Error GoTo TestFail
    
    'Arrange:
    Dim tag As New clsTag
    Dim ws As Worksheet
    Dim StartCell As Range
    Set tag = New clsTag
    Set ws = ThisWorkbook.Worksheets("TestFailureCodeTemplate")
    tag.ID = "TEST-TAG"
    tag.FailureCode = "TestFailureCodeTemplate"
    tag.Criticality = "x"
    
    'Act:
    tag.SetDefaultCriticalityByFailureCode
    
    'Assert:
    Assert.Istrue (tag.Criticality = "C")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub


