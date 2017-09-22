Attribute VB_Name = "TestLoadingTags"
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
Public Sub TestReadLotsOfTags()
    On Error GoTo TestFail

    Dim tag As clsTag
    Dim tags As clsTags
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tagArray As Variant
    Dim x As Long          ' >32000 entries

    'Act:
    Set wb = Application.Workbooks("WND Criticality Template.xlsx")
    Set ws = wb.Worksheets("AssetRegisterDefaultCodeApplied")
    Set tag = New clsTag
    Set tags = New clsTags
    Set tbl = ws.ListObjects("AssetRegisterTbl")
    Debug.Print tbl.Name
    'Create Array List from Table
    tagArray = tbl.DataBodyRange
'    'Loop through each item in Third Column of Table (displayed in Immediate Window [ctrl + g])
'    For x = LBound(tagArray) To UBound(tagArray)
'        tag.TagID = tagArray(x, 1)
'        tag.TagDescription = tagArray(x, 2)
'        With tag
'            'Debug.Print x, .TagID, .TagDescription
'        End With
'    Next x
    'tags.LoadTable tbl 'out of memory error with 100k rows
    tags.LoadArray tagArray
    
    'Assert:
    'Debug.Print x, UBound(tagArray), LBound(tagArray)
    'Debug.Print tag.TagID, tag.TagDescription
    'Assert.IsTrue (tag.TagID = "BP")
    Debug.Print UBound(tagArray)
    Debug.Print tags.Count
    Debug.Print "Tag 1 = ", tags.Item(1).ID
    Assert.istrue (tags.Count = 104572)  'NB - runs over end of table...
    'Assert.IsTrue (UBound(tagArray) - LBound(tagArray) + 1 = 104573) 'Appears to default to 1 indexing

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub

'@TestMethod
Public Sub TestFullTags()
    On Error GoTo TestFail

    Dim tag As clsTag
    Dim tags As clsTags
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tagArray As Variant
    Dim x As Long          ' >32000 entries

    'Act:
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("TestFullTags")
    Set tag = New clsTag
    Set tags = New clsTags
    Set tbl = ws.ListObjects("TestFullTags")
    Debug.Print tbl.Name
    'Create Array List from Table
    tagArray = tbl.DataBodyRange
    tags.LoadArray tagArray
    tag.FillTagHeaders ThisWorkbook.Worksheets("Output").Range("B10")
    tags.OutputTagListings ThisWorkbook.Worksheets("Output").Range("B11")
    
    'Assert:
    'Debug.Print x, UBound(tagArray), LBound(tagArray)
    'Debug.Print tag.TagID, tag.TagDescription
    'Assert.IsTrue (tag.TagID = "BP")
    'Debug.Print UBound(tagArray)
    Debug.Print tags.Count
    Debug.Print "Tag 1 = ", tags.Item(1).ID
    Assert.istrue (tags.Count = 23)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub

