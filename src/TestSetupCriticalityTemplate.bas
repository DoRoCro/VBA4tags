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
    Dim codesTable As ListObject 'this is the failure codes table
    Dim codeRow As ListRow       'this is the row in the falure codes table
    Dim defaultsTable As ListObject  'this is the table with default values for a failure code
    
    Set codesTable = ThisWorkbook.Worksheets("TestDefaultCriticalities").ListObjects("TestFailureCodeDefaultCriticalitiesTable")
    Set codeRow = codesTable.ListRows(1)
    
    'Set initial value to something that shouldn't be there at end
    'Safety
    ThisWorkbook.Worksheets("TestFailureCodeTemplate").Range("B16").Formula = "AA"
    ThisWorkbook.Worksheets("TestFailureCodeTemplate").Range("C16").Value = 9
    'Env
    ThisWorkbook.Worksheets("TestFailureCodeTemplate").Range("B22").Formula = "AA"
    ThisWorkbook.Worksheets("TestFailureCodeTemplate").Range("C22").Value = 9
    'Prod
    ThisWorkbook.Worksheets("TestFailureCodeTemplate").Range("B28").Formula = "AA"
    ThisWorkbook.Worksheets("TestFailureCodeTemplate").Range("C28").Value = 9
    'Non-Financial business
    ThisWorkbook.Worksheets("TestFailureCodeTemplate").Range("B34").Formula = "AA"
    ThisWorkbook.Worksheets("TestFailureCodeTemplate").Range("C34").Value = 9
    'using the same table under two names here, this one need the lookup as it might be different order to above
    Set defaultsTable = ThisWorkbook.Worksheets("TestDefaultCriticalities").ListObjects("TestFailureCodeDefaultCriticalitiesTable")
    
    'Act:
    Call CopyDefaultCriticalitiesIntoTemplateWorksheet(codeRow, ThisWorkbook.Worksheets("TestFailureCodeTemplate"), defaultsTable)
    'Assert:
    Debug.Print ThisWorkbook.Worksheets("TestDefaultCriticalities").Range("B1").Formula
    Assert.istrue ThisWorkbook.Worksheets("TestFailureCodeTemplate").Range("B1").Formula = "FA_CFBC"
    Assert.istrue ThisWorkbook.Worksheets("TestFailureCodeTemplate").Range("B16").Formula = "E"
    Assert.istrue ThisWorkbook.Worksheets("TestFailureCodeTemplate").Range("C16").Value = 8
    Assert.istrue ThisWorkbook.Worksheets("TestFailureCodeTemplate").Range("B22").Formula = "H"
    Assert.istrue ThisWorkbook.Worksheets("TestFailureCodeTemplate").Range("C22").Formula = "2"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


