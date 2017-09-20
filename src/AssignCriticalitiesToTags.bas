Attribute VB_Name = "AssignCriticalitiesToTags"
'@Folder("CriticalityAssignment")

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
    Set tags = New clsTags
    Set Systems = New clsSystems
    Set disciplines = New clsDisciplines
    Call LoadTables(tags, Systems, disciplines)
    Debug.Print "Tags count ="; tags.Count
    Debug.Print "Systems count ="; Systems.Count
    Debug.Print "Disciplines count ="; disciplines.Count

'foreach tag
    'lookup failure code output
    ' Set criticalities by failure code as first pass
    Dim tag As clsTag
    For Each tag In tags.All
        'Debug.Print tag.ID
        Select Case tag.Status
            Case "DEL"
                tag.Criticality = "D"
            Case "SOFT"
                tag.Criticality = "S"
            Case Else
                Call SetTagCriticalityByFailureCode(tag)
        End Select
        
        
    Next
    Debug.Print "Default criticalities assigned"

    'set criticality using default MAH barrier, if defined
    'if isUtility then look at downgrade options / revising MAH barrier
    'if isSIL then set as LOPA/IPL in Non-fin business, which will give criticality A
    'if isSIS then set as LOPA/IPL in Non-fin business, which will give criticality A
    
    
    'copy results of template to row for tag into discipline workbook with comments
    '  / justification

'endforeach
End Sub


Sub LoadTables(tags As clsTags, Systems As clsSystems, Discplines As clsDisciplines)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tagsArray As Variant
    Set wb = Workbooks(CriticalityWbName)
    Set ws = wb.Worksheets(tagsWorksheetName)
    Set tbl = ws.ListObjects(tagsTableName)

'Read in tags
    tagsArray = tbl.DataBodyRange
    tags.LoadArray tagsArray
    Debug.Print "finished loading tags, count ="; tags.Count
    Set tagsArray = Nothing
    
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
    Select Case tag.FailureCode
        Case "SOFT", "LOOP", vbNullString
        'If tag.FailureCode <> "" Then
            'Debug.Print "Blank Failure code for "; tag.ID
            tag.Criticality = "F"
        'Else
            ' some error message here
        Case Else
            Set ws = wb.Worksheets(tag.FailureCode)
            tag.Criticality = ws.Range("K1")
    End Select
        'End If
    'ws.Activate
    
End Sub

