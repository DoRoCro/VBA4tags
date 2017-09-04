Attribute VB_Name = "TestTags"
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
Public Sub TestGetLetTagID()
    On Error GoTo TestFail
    
    'Arrange:
    Dim tag As clsTag
    Const strTag = "TEST-TAG"
    Set tag = New clsTag
    'Act:
    tag.TagID = "TEST-TAG"
    'Assert:
    Assert.IsTrue ("TEST-TAG" = tag.TagID)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetLetTagDescription()
    On Error GoTo TestFail
    
    'Arrange:
    Dim tag As clsTag
    Const strTagDesc = "TEST-TAG-DESC"
    Set tag = New clsTag
    'Act:
    tag.TagDescription = strTagDesc
    'Assert:
    Assert.IsTrue ("TEST-TAG-DESC" = tag.TagDescription)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TestGetTagIDFromCell()
    On Error GoTo TestFail
    
    'Arrange:
    Dim tag As clsTag
    Dim wb As Workbook
    Dim ws As Worksheet
    'Act:
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TestingData")
    Set tag = New clsTag
    tag.TagID = ws.Cells(2, 1).Value
    Debug.Print tag.TagID
    'Assert:
    Assert.IsTrue (tag.TagID = "AB12345A")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod
Public Sub TestGetTagDescFromCell()
    On Error GoTo TestFail
    
    'Arrange:
    Dim tag As clsTag
    Dim wb As Workbook
    Dim ws As Worksheet
    'Act:
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TestingData")
    Set tag = New clsTag
    tag.TagID = ws.Cells(2, 2).Value
    Debug.Print tag.TagID
    'Assert:
    Assert.IsTrue (tag.TagID = "A TAG FOR TESTING")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod
Public Sub TestGetTagFromTable()
    On Error GoTo TestFail
    
    'Arrange:
    Dim tag As clsTag
    Dim tags As Collection
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim x As Integer
    

    'Act:
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TestingData")
    Set tag = New clsTag
    Set tags = New Collection
    Set tbl = ws.ListObjects("TagMinimal")
    Debug.Print tbl.Name
    
    'Loop Through Every Row in Table
    For x = 1 To tbl.Range.Rows.Count
        tag.TagID = tbl.Range.Rows(x).Columns(1)
        tags.Add tag
        Debug.Print tag.TagID
    Next x

    'Assert:
    Assert.IsTrue (tbl.Name = "TagMinimal")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


