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
    Dim Tag As clsTag
    Const strTag As String = "TEST-TAG"
    Set Tag = New clsTag
    'Act:
    Tag.ID = strTag
    'Assert:
    Assert.isTrue (strTag = Tag.ID)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestGetLetTagDescription()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Tag As clsTag
    Const strTagDesc As String = "TEST-TAG-DESC"
    Set Tag = New clsTag
    'Act:
    Tag.Description = strTagDesc
    'Assert:
    Assert.isTrue ("TEST-TAG-DESC" = Tag.Description)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TestGetTagIDFromCell()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Tag As clsTag
    Dim wb As Workbook
    Dim ws As Worksheet
    'Act:
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TestingData")
    Set Tag = New clsTag
    Tag.ID = ws.Cells(2, 1).Value
    Debug.Print Tag.ID
    'Assert:
    Assert.isTrue (Tag.ID = "AB12345A")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod
Public Sub TestGetTagDescFromCell()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Tag As clsTag
    Dim wb As Workbook
    Dim ws As Worksheet
    'Act:
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TestingData")
    Set Tag = New clsTag
    Tag.ID = ws.Cells(2, 2).Value
    Debug.Print Tag.ID
    'Assert:
    Assert.isTrue (Tag.ID = "A TAG FOR TESTING")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
'@TestMethod
Public Sub TestGetTagFromTable()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Tag As clsTag
    Dim Tags As Collection
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim x As Integer
    

    'Act:
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TestingData")
    Set Tag = New clsTag
    Set Tags = New Collection
    Set tbl = ws.ListObjects("TagMinimal")
    Debug.Print tbl.Name
    
    'Loop Through Every Row in Table
    For x = 1 To tbl.Range.Rows.Count
        Tag.ID = tbl.Range.Rows(x).Columns(1)
        Tags.Add Tag
        Debug.Print Tag.ID
    Next x

    'Assert:
    Assert.isTrue (tbl.Name = "TagMinimal")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestReadTableCreateTags()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Tag As clsTag
    Dim Tags As Collection
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tagArray As Variant
    Dim x As Integer
    
    'Act:
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TestingData")
    Set Tag = New clsTag
    Set Tags = New Collection
    Set tbl = ws.ListObjects("TagMinimal")
    Debug.Print tbl.Name
    'Create Array List from Table
    tagArray = tbl.DataBodyRange
    'Loop through each item in Third Column of Table (displayed in Immediate Window [ctrl + g])
    For x = LBound(tagArray) To UBound(tagArray)
        Tag.ID = tagArray(x, 1)
        Tag.Description = tagArray(x, 2)
        With Tag
            Debug.Print x, .ID, .Description
        End With
    Next x
    'Assert:
    Debug.Print x, UBound(tagArray), LBound(tagArray)
    Assert.isTrue (Tag.ID = "E-K-2421")
    Assert.isTrue (x = 3)  'NB - runs over end of table...
    Assert.isTrue (UBound(tagArray) - LBound(tagArray) + 1 = 2) 'Appears to default to 1 indexing

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

'@TestMethod
Public Sub TestLoadTagsFromTable()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Tags As clsTags
    Set Tags = New clsTags
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("TestTags").ListObjects("TestTagsTable")
    
    'Act:
    Tags.LoadTable tbl
    Debug.Print Tags.Item(1).ID
    'Assert:
    Assert.isTrue (Tags.Item(1).ID = "E-VG-29-069")
    Assert.isTrue (Tags.Item(2).FailureCode = "FA_PEVB")
    Assert.isTrue Tags.Item(5).isSIS
    Assert.isFalse Tags.Item(5).isSIL
    Assert.isTrue Tags.Item(6).isSIL
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub




