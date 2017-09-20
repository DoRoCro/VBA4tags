Attribute VB_Name = "TestDisciplineClass"
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
Public Sub TestDisciplineLoadsFromTable()
    On Error GoTo TestFail
    

    Dim disciplines As Collection
    Dim discipline As clsDiscipline
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim x As Integer
    

    'Act:
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TestDisciplines")
    Set disciplines = New Collection
    Set tbl = ws.ListObjects("DisciplinesList")
    Debug.Print tbl.Name
    
    'Loop Through Every Row in Table  NB row 1 is headers
    For x = 2 To tbl.Range.Rows.Count
        Set discipline = New clsDiscipline
        discipline.ID = tbl.Range.Rows(x).Columns(1)
        disciplines.Add discipline
        Debug.Print discipline.ID
    Next x
    Debug.Print disciplines(2).ID
    'Assert:
    Assert.istrue (tbl.Name = "DisciplinesList")
    Assert.istrue (disciplines(2).ID = "ELEC")
    Assert.istrue (disciplines(6).ID = "MECH")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestDisciplinesClassLoads()
    On Error GoTo TestFail
    
    Dim disciplines As clsDisciplines
    Dim discipline As clsDiscipline
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim x As Integer
    

    'Act:
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TestDisciplines")
    Set disciplines = New clsDisciplines
    Set tbl = ws.ListObjects("DisciplinesList")
    Debug.Print tbl.Name
    
    'load as a table
    disciplines.LoadTable tbl

    Debug.Print disciplines.Item(2).ID
    'Assert:
    Assert.istrue (tbl.Name = "DisciplinesList")
    Assert.istrue (disciplines.Item(2).ID = "ELEC")
    Assert.istrue (disciplines.Item(6).ID = "MECH")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCreateOutputSheet()
    On Error GoTo TestFail
    
    'Arrange:
    Dim discipline As clsDiscipline
    Set discipline = New clsDiscipline
    Dim ws As Worksheet
    discipline.ID = "TEST"
    'Act:
    Set ws = discipline.CreateDisciplineOutputSheet(ThisWorkbook.Sheets)
    'Assert:
    Assert.istrue (ws.Name = discipline.ID)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestCreateOutputSheetsAllDisciplines()
    On Error GoTo TestFail
    
    'Arrange:
    Dim disciplines As clsDisciplines
    Dim ws As Worksheet
    Set disciplines = New clsDisciplines
    disciplines.LoadTable ThisWorkbook.Worksheets("TestDisciplines").ListObjects("DisciplinesList")
    
    'Act:
    disciplines.createOutputSheetsByDiscipline
    'Assert:
    Assert.istrue (ThisWorkbook.Worksheets(Sheets.Count).Name = "TELE")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


