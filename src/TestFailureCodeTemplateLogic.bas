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
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
     'Reset all values in template to H1 before starting
    Dim rowStr As Variant
    For Each rowStr In Array("16", "22", "28", "34")
        ws.Range("B" & rowStr).Value = "A"
        ws.Range("C" & rowStr).Value = "8"
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
    
    'Reset all values in template to H1 before starting
    Dim rowStr As Variant
    For Each rowStr In Array("16", "22", "28", "34")
        ws.Range("B" & rowStr).Value = "H"
        ws.Range("C" & rowStr).Value = "1"
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
    Assert.istrue (ws.Range("K1") = "C")
    Assert.istrue (ws.Range("G6").Value = "NO") ' not SCE

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestSafetyA1GivesCAndSCE()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    ws.Range("B16").Formula = "A"
    'Assert:
    Assert.istrue (ws.Range("K1") = "C")
    Assert.istrue (ws.Range("G6").Value = "YES")   ' is SCE
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestSafetyA7GivesCAndSCE()
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    ws.Range("B16").Formula = "A"
    ws.Range("C16").Formula = "7"
    
    'Assert:
    Assert.istrue (ws.Range("K1") = "C")        ' no rule explicitly sets high criticality from safety impact, but expect MAH override in practice
    Assert.istrue (ws.Range("G6").Value = "YES")   ' is SCE
    Assert.istrue (ws.Range("G8").Value = 10)   ' CMMS Location priority
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

