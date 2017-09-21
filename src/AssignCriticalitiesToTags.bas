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
Const MAHSheetName = "MAHBarrierSetup"
Const MAHTableName = "MAHBarrierForFailureCode"

Private tags As clsTags
Private disciplines As clsDisciplines
Private Systems As clsSystems
Private MAHprocess As clsMAHlist
Private MAHutility As clsMAHlist

Sub AssignCriticalities()
    Dim tag As clsTag
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

    'set criticality using default MAH barrier, if defined
    'if isUtility then look at downgrade options / revising MAH barrier
    'if isSIL then set as LOPA/IPL in Non-fin business, which will give criticality A
    'if isSIS then set as LOPA/IPL in Non-fin business, which will give criticality A
    
    
    'copy results of template to row for tag into discipline workbook with comments
    '  / justification
    tags.Item(1).FillTagHeaders ThisWorkbook.Worksheets("Output").Range("B50")
    Set tags = tags.RemoveStatus("DEL")
    Set tags = tags.RemoveStatus("SOFT")
    Set tags = tags.RemoveStatus(vbNullString)
    Set tags = tags.RemoveStatus("DRAFT")
    
'    Dim pipingTags As clsTags
'    Set pipingTags = New clsTags
'    Set pipingTags = tags.byDiscipline(disciplines.Item(8))
'
'    pipingTags.OutputTagListings ThisWorkbook.Worksheets("Output").Range("B51")
    
    Dim disciplineTags As clsTags
    Set disciplineTags = New clsTags
    Dim discWs As Worksheet
    'disciplines.createOutputSheetsByDiscipline
    Dim Discipline As clsDiscipline
    Dim counter As Long
    counter = 0
    For Each Discipline In disciplines.All
        Set disciplineTags = tags.byDiscipline(Discipline)
        Excel.Application.ScreenUpdating = False
        For Each tag In disciplineTags.All
            counter = counter + 1
            'Debug.Print tag.ID
            Select Case tag.Status
                Case "DEL"
                    tag.Criticality = "D"
                Case "SOFT"
                    tag.Criticality = "S"
                Case Else
                    If Systems.Contains(tag.SystemID) Then
                        If Systems.FindByNumber(tag.SystemID).isUtility Then
                            Call SetTagCriticalityByFailureCode(tag, MAHutility)
                        Else
                            Call SetTagCriticalityByFailureCode(tag, MAHprocess)
                        End If
                    Else
                        tag.Criticality = "X"
                    End If
            End Select
            If counter Mod 100 = 0 Then Debug.Print "Tags processed = ", counter
        Next
        Excel.Application.ScreenUpdating = True
        Debug.Print "Default criticalities assigned for ", Discipline.ID
        
        Set discWs = Discipline.CreateDisciplineOutputSheet(ThisWorkbook.Sheets)
'        Select Case Discipline.ID
'            Case vbNullString
'                discWs = Worksheets("BLANKS")
'            Case "N/A"
'                discWs = Worksheets("N_A")
'            Case Else
'                discWs = Worksheets(Discipline.ID)
'        End Select
        disciplineTags.OutputTagListings discWs.Range("A1")
    Next Discipline
    
    
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
    Set ws = wb.Worksheets(MAHSheetName)
    Set tbl = ws.ListObjects(MAHTableName)
    Set MAHprocess = New clsMAHlist
    Set MAHutility = New clsMAHlist
    MAHprocess.LoadTableProcess tbl
    MAHutility.LoadTableUtility tbl
    
'Read in Ex list - optional, if not added to tag.


End Sub

'Sub CalculateTagCriticality(tag As clsTag, Systems As clsSystems, Optional MAHBarriers As Collection, Optional Overrides As Collection)
Sub SetTagCriticalityByFailureCode(tag As clsTag, MAH As clsMAHlist)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim MAHCell As Range
    Dim resetComponent As String
    Dim resetComment As String
    Set wb = Workbooks(CriticalityWbName)
    Select Case tag.FailureCode
        Case "SOFT", "LOOP", vbNullString      'ignore these failure codes
            tag.Criticality = "F"
        Case Else
            Set ws = wb.Worksheets(tag.FailureCode)
            Set MAHCell = ws.Range("I17")
            resetComponent = MAHCell.Text
            resetComment = ws.Range("I19").Text
            If MAH.Count > 0 Then                              'TODO think how to refactor this or change parameters to function
                MAHCell.Value = MAH.FindByID(tag.FailureCode).Component
            End If
            tag.Criticality = ws.Range("K1")
            tag.RiskImpact(Safety) = ws.Range("B9")
            tag.RiskImpact(Environment) = ws.Range("B10")
            tag.RiskImpact(Production) = ws.Range("B11")
            tag.RiskImpact(Business) = ws.Range("B12")
            tag.RiskLikelihood(Safety) = ws.Range("C9")
            tag.RiskLikelihood(Environment) = ws.Range("C10")
            tag.RiskLikelihood(Production) = ws.Range("C11")
            tag.RiskLikelihood(Business) = ws.Range("C12")
            tag.Justification = "MAH Barrier: " & ws.Range("I17").Text & "; " & _
                                "MAH comment: " & ws.Range("I19").Text & "; " & _
                                "IPL/LOPA = " & ws.Range("J35").Text & _
                                "; Regulatory override = " & ws.Range("I37").Text
            MAHCell.Value = resetComponent
            ws.Range("i19") = resetComment
    End Select
        'End If
    'ws.Activate
    
End Sub


