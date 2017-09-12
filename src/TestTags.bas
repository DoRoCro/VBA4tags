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
        tag.TagID = tagArray(x, 1)
        tag.TagDescription = tagArray(x, 2)
        With tag
            Debug.Print x, .TagID, .TagDescription
        End With
    Next x
    'Assert:
    Debug.Print x, UBound(tagArray), LBound(tagArray)
    Assert.IsTrue (tag.TagID = "E-K-2421")
    Assert.IsTrue (x = 3)  'NB - runs over end of table...
    Assert.IsTrue (UBound(tagArray) - LBound(tagArray) + 1 = 2) 'Appears to default to 1 indexing

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
'Public Sub TestReadLotsOfTags()
'    On Error GoTo TestFail
'
'    Dim tag As clsTag
'    Dim tags As Collection
'    Dim wb As Workbook
'    Dim ws As Worksheet
'    Dim tbl As ListObject
'    Dim tagArray As Variant
'    Dim x As Long          ' >32000 entries
'
'    'Act:
'    Set wb = Application.Workbooks("Criticality Assignment Spreadsheet 2017-09-04.xlsx")
'    Set ws = wb.Worksheets("AssetRegisterTbl")
'    Set tag = New clsTag
'    Set tags = New Collection
'    Set tbl = ws.ListObjects("AssetRegisterTbl")
'    Debug.Print tbl.Name
'    'Create Array List from Table
'    tagArray = tbl.DataBodyRange
'    'Loop through each item in Third Column of Table (displayed in Immediate Window [ctrl + g])
'    For x = LBound(tagArray) To UBound(tagArray)
'        tag.TagID = tagArray(x, 1)
'        tag.TagDescription = tagArray(x, 2)
'        With tag
'            'Debug.Print x, .TagID, .TagDescription
'        End With
'    Next x
'    'Assert:
'    Debug.Print x, UBound(tagArray), LBound(tagArray)
'    Debug.Print tag.TagID, tag.TagDescription
'    Assert.IsTrue (tag.TagID = "BP")
'    Assert.IsTrue (x = 104574)  'NB - runs over end of table...
'    Assert.IsTrue (UBound(tagArray) - LBound(tagArray) + 1 = 104573) 'Appears to default to 1 indexing
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub


