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
    

    Dim Disciplines As Collection
    Dim Discipline As clsDiscipline
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim x As Integer
    

    'Act:
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TestDisciplines")
    Set Disciplines = New Collection
    Set tbl = ws.ListObjects("DisciplinesList")
    Debug.Print tbl.Name
    
    'Loop Through Every Row in Table  NB row 1 is headers
    For x = 2 To tbl.Range.Rows.Count
        Set Discipline = New clsDiscipline
        Discipline.ID = tbl.Range.Rows(x).Columns(1)
        Disciplines.Add Discipline
        Debug.Print Discipline.ID
    Next x
    Debug.Print Disciplines(2).ID
    'Assert:
    Assert.IsTrue (tbl.Name = "DisciplinesList")
    Assert.IsTrue (Disciplines(2).ID = "ELEC")
    Assert.IsTrue (Disciplines(6).ID = "MECH")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestDisciplinesClassLoads()
    On Error GoTo TestFail
    
    Dim Disciplines As clsDisciplines
    Dim Discipline As clsDiscipline
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim x As Integer
    

    'Act:
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TestDisciplines")
    Set Disciplines = New clsDisciplines
    Set tbl = ws.ListObjects("DisciplinesList")
    Debug.Print tbl.Name
    
    'load as a table
    Disciplines.LoadTable tbl

    Debug.Print Disciplines.Item(2).ID
    'Assert:
    Assert.IsTrue (tbl.Name = "DisciplinesList")
    Assert.IsTrue (Disciplines.Item(2).ID = "ELEC")
    Assert.IsTrue (Disciplines.Item(6).ID = "MECH")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


