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
    Const strTag As String = "TEST-TAG"
    Set tag = New clsTag
    'Act:
    tag.ID = strTag
    'Assert:
    Assert.IsTrue (strTag = tag.ID)

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
    Const strTagDesc As String = "TEST-TAG-DESC"
    Set tag = New clsTag
    'Act:
    tag.Description = strTagDesc
    'Assert:
    Assert.IsTrue ("TEST-TAG-DESC" = tag.Description)

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
    tag.ID = ws.Cells(2, 1).Value
    Debug.Print tag.ID
    'Assert:
    Assert.IsTrue (tag.ID = "AB12345A")

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
    tag.ID = ws.Cells(2, 2).Value
    Debug.Print tag.ID
    'Assert:
    Assert.IsTrue (tag.ID = "A TAG FOR TESTING")

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
        tag.ID = tbl.Range.Rows(x).Columns(1)
        tags.Add tag
        Debug.Print tag.ID
    Next x

    'Assert:
    Assert.IsTrue (tbl.Name = "TagMinimal")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestReadTableCreateTags()
    On Error GoTo TestFail
    
    'Arrange:
    Dim tag As clsTag
    Dim tags As Collection
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tagArray As Variant
    Dim x As Integer
    
    'Act:
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TestingData")
    Set tag = New clsTag
    Set tags = New Collection
    Set tbl = ws.ListObjects("TagMinimal")
    Debug.Print tbl.Name
    'Create Array List from Table
    tagArray = tbl.DataBodyRange
    'Loop through each item in Third Column of Table (displayed in Immediate Window [ctrl + g])
    For x = LBound(tagArray) To UBound(tagArray)
        tag.ID = tagArray(x, 1)
        tag.Description = tagArray(x, 2)
        With tag
            Debug.Print x, .ID, .Description
        End With
    Next x
    'Assert:
    Debug.Print x, UBound(tagArray), LBound(tagArray)
    Assert.IsTrue (tag.ID = "E-K-2421")
    Assert.IsTrue (x = 3)  'NB - runs over end of table...
    Assert.IsTrue (UBound(tagArray) - LBound(tagArray) + 1 = 2) 'Appears to default to 1 indexing

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TestLoadTagsFromTable()
    On Error GoTo TestFail
    
    'Arrange:
    Dim tags As clsTags
    Set tags = New clsTags
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("TestTags").ListObjects("TestTagsTable")
    
    'Act:
    tags.LoadTable tbl
    Debug.Print tags.Item(1).ID
    'Assert:
    Assert.IsTrue (tags.Item(1).ID = "E-VG-29-069")
    Assert.IsTrue (tags.Item(2).FailureCode = "FA_PEVB")
    Assert.IsTrue tags.Item(5).isSIS
    Assert.isFalse tags.Item(5).isSIL
    Assert.IsTrue tags.Item(6).isSIL
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub




