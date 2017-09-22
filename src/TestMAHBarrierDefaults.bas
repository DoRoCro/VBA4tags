Attribute VB_Name = "TestMAHBarrierDefaults"
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
Public Sub TestMAHbasics()
    On Error GoTo TestFail
    
    'Arrange:
    Dim MAH As clsMAHDefault
    Set MAH = New clsMAHDefault
    'Act:
    MAH.ID = "FA_CFBC"
    MAH.Component = "Some text here"
    MAH.Family = "Family"
    MAH.Comment = "Comment"
    MAH.TypCriticality = "D"

    'Assert:
    Assert.AreEqual MAH.ID, "FA_CFBC"
    Assert.AreEqual MAH.TypCriticality, "D"
    Assert.AreEqual MAH.Family, "Family"
    Assert.AreEqual MAH.Comment, "Comment"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub

'@TestMethod
Public Sub TestMAHAddtoList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim MAH As clsMAHDefault
    Dim MAHlist As clsMAHlist
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set MAH = New clsMAHDefault
    Set MAHlist = New clsMAHlist
    MAH.ID = "FA_CFBC"
    MAH.Component = "Some text here"
    MAH.Family = "Family"
    MAH.Comment = "Comment"
    MAH.TypCriticality = "D"

    'Act:
    MAHlist.Add MAH
    'Assert:
    Assert.istrue (MAHlist.Item(1).ID = "FA_CFBC")
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub

'@TestMethod
Public Sub TestLoadMAHDefaults()
    On Error GoTo TestFail
    
    'Arrange:
    Dim MAH As clsMAHDefault
    Dim MAHlist As clsMAHlist
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set wb = Application.Workbooks("WND Criticality Template.xlsx")
    Set ws = wb.Worksheets("MAHBarrierSetup")
    Set MAH = New clsMAHDefault
    Set MAHlist = New clsMAHlist
    Set tbl = ws.ListObjects("MAHBarrierForFailureCode")
    Debug.Print tbl.Name
    
    'Act:
    MAHlist.LoadTableUtility tbl


    'Assert:
    Assert.istrue (MAHlist.Item(1).ID = "FA_CFBC")    'based on initial dataset
    Assert.istrue (MAHlist.Item(2).Family = "#N/A")   'cope with error cells as test
    Assert.istrue (MAHlist.Item(2).TypCriticality = "#")   'stores only first character in cell
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub

'@TestMethod
Public Sub TestMAHFindByFailureCode()
    On Error GoTo TestFail
    
    'Arrange:
    Dim MAH As clsMAHDefault
    Dim MAHlist As clsMAHlist
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    Set wb = Application.Workbooks("WND Criticality Template.xlsx")
    Set ws = wb.Worksheets("MAHBarrierSetup")
    Set MAH = New clsMAHDefault
    Set MAHlist = New clsMAHlist
    Set tbl = ws.ListObjects("MAHBarrierForFailureCode")
    Debug.Print tbl.Name
    MAHlist.LoadTableUtility tbl
    
    'Act:
    

    'Assert:
    Assert.istrue (MAHlist.FindByID("FA_CFBC").ID = "FA_CFBC")
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub
