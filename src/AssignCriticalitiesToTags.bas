Attribute VB_Name = "AssignCriticalitiesToTags"
'@Folder("VBAProject")

Option Explicit

Const CriticalityWbName As String = "WND Criticality Template.xlsx"
Const tagsTableName = "AssetRegisterTbl"
Const tagsWorksheetName = "AssetRegisterDefaultCodeApplied"
Const DisciplinesSheetName = "DataTables"
Const DisciplinesTableName = "DisciplinesList"
Const SystemsSheetName = "SystemsUtilities"
Const SystemsTableName = "SystemsList"

Private tags As clsTags
Private disciplines As clsDisciplines
Private Systems As clsSystems

Sub AssignCriticalities()
    Call LoadTags
    Debug.Print "Tags count ="; tags.Count
    Debug.Print "Systems count ="; Systems.Count
    Debug.Print "Disciplines count ="; disciplines.Count

'foreach tag
    'lookup failure code output
    'set criticality using default MAH barrier, if defined
    'if isUtility then look at downgrade options / revising MAH barrier
    'if isSIL then set as LOPA/IPL in Non-fin business, which will give criticality A
    'if isSIS then set as LOPA/IPL in Non-fin business, which will give criticality A
    
    
    'copy results of template to row for tag into discipline workbook with comments
    '  / justification

'endforeach
End Sub


Sub LoadTags()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tagsArray As Variant
    Set tags = New clsTags
    Set wb = Workbooks(CriticalityWbName)
    Set ws = wb.Worksheets(tagsWorksheetName)
    Set tbl = ws.ListObjects(tagsTableName)

'Read in tags
    tagsArray = tbl.DataBodyRange
    tags.LoadArray tagsArray
    Debug.Print "finished loading tags, count ="; tags.Count

'Read in disciplines
    Set ws = wb.Worksheets(DisciplinesSheetName)
    Set tbl = ws.ListObjects(DisciplinesTableName)
    Set disciplines = New clsDisciplines
    disciplines.LoadTable tbl
    Debug.Print "finished loading disciplines, count ="; disciplines.Count

'Read in systems
    Set ws = wb.Worksheets(SystemsSheetName)
    Set tbl = ws.ListObjects(SystemsTableName)
    Set Systems = New clsSystems
    Systems.LoadTable tbl
    Debug.Print "finished loading systems, count ="; Systems.Count

'Read in MAH
'Read in Ex list - optional, if not added to tag.


End Sub

'Sub CalculateTagCriticality(tag As clsTag, Systems As clsSystems, Optional MAHBarriers As Collection, Optional Overrides As Collection)
Sub SetTagCriticalityByFailureCode(tag As clsTag)
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Workbooks(CriticalityWbName)
    If tag.FailureCode <> "" Then
        Set ws = wb.Worksheets(tag.FailureCode)
        tag.Criticality = ws.Range("K1")
    Else
        ' some error message here
    End If
    'ws.Activate
    
End Sub

