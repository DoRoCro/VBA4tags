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
    Dim Discipline As clsDiscipline
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
        Set Discipline = New clsDiscipline
        Discipline.ID = tbl.Range.Rows(x).Columns(1)
        disciplines.Add Discipline
        Debug.Print Discipline.ID
    Next x
    Debug.Print disciplines(2).ID
    'Assert:
    Assert.isTrue (tbl.Name = "DisciplinesList")
    Assert.isTrue (disciplines(2).ID = "ELEC")
    Assert.isTrue (disciplines(6).ID = "MECH")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub

'@TestMethod
Public Sub TestDisciplinesClassLoads()
    On Error GoTo TestFail
    
    Dim disciplines As clsDisciplines
    Dim Discipline As clsDiscipline
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
    Assert.isTrue (tbl.Name = "DisciplinesList")
    Assert.isTrue (disciplines.Item(2).ID = "ELEC")
    Assert.isTrue (disciplines.Item(6).ID = "MECH")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub

'@TestMethod
Public Sub TestCreateOutputSheet()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Discipline As clsDiscipline
    Set Discipline = New clsDiscipline
    Dim ws As Worksheet
    Discipline.ID = "TEST"
    'Act:
    Set ws = Discipline.CreateDisciplineOutputSheet(ThisWorkbook.Sheets)
    'Assert:
    Assert.isTrue (ws.Name = Discipline.ID)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
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
    Assert.isTrue (ThisWorkbook.Worksheets(Sheets.Count).Name = "TELE")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub


