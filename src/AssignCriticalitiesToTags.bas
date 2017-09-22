Attribute VB_Name = "AssignCriticalitiesToTags"
'@Folder("CriticalityAssignment")

Option Explicit


Const tagsTableName = "AssetRegisterTbl"
Const tagsWorksheetName = "AssetRegisterDefaultCodeApplied"
Const DisciplinesSheetName = "DataTables"
Const DisciplinesTableName = "DisciplinesList"
Const SystemsSheetName = "SystemsUtilities"
Const SystemsTableName = "SystemsList"
Const MAHSheetName = "MAHBarrierSetup"
Const MAHTableName = "MAHBarrierForFailureCode"
Const FailureCodesSheetName = "FailureCodes"
Const FailureCodesTableName = "ASSET_C_FailureCodesList"

Public CriticalityWbName As String
Private tags As clsTags
Private disciplines As clsDisciplines
Private Systems As clsSystems
Private MAHprocess As clsMAHlist
Private MAHutility As clsMAHlist
Private FailureCodes As clsFailureCodes


Sub AssignCriticalities()
    Dim tag As clsTag
    Dim disciplineTags As clsTags
    Dim discWs As Worksheet
    Dim Discipline As clsDiscipline
    Dim counter As Long
    
    CriticalityWbName = "WND Criticality Template.xlsx"
    
    Set tags = New clsTags
    Set Systems = New clsSystems
    Set disciplines = New clsDisciplines
    Set FailureCodes = New clsFailureCodes
    Call LoadTables(tags, Systems, disciplines)
    Debug.Print "Tags count ="; tags.Count
    Debug.Print "Systems count ="; Systems.Count
    Debug.Print "Disciplines count ="; disciplines.Count
    Debug.Print "Systems count ="; Systems.Count
    Debug.Print "Failurecodes count ="; FailureCodes.Count
    Debug.Print "MAHprocess count ="; MAHprocess.Count
    Debug.Print "MAHutility count ="; MAHutility.Count
    
    'filter out unwanted tags by STATUS code
    Set tags = tags.RemoveStatus("DEL")
    Set tags = tags.RemoveStatus("SOFT")
    Set tags = tags.RemoveStatus(vbNullString)
    Set tags = tags.RemoveStatus("DRAFT")
    
    
    
    Set disciplineTags = New clsTags
    counter = 0
    For Each Discipline In disciplines.All
        Set disciplineTags = tags.byDiscipline(Discipline)
        'Excel.Application.ScreenUpdating = False
        'setup for process tags
        FailureCodes.SetupForSystemGroup MAHprocess
        
        'assign process tags
        disciplineTags.ProcessTags(Systems).AssignDefaultCriticalities
        
        'setup for utility tags
        FailureCodes.SetupForSystemGroup MAHutility
        
        'assign utility tags
        disciplineTags.UtilityTags(Systems).AssignDefaultCriticalities
        
        'deal with nosystem tags, as utility for now
        disciplineTags.NoSystemTags(Systems).AssignDefaultCriticalities

        'Excel.Application.ScreenUpdating = True
        Debug.Print "Default criticalities assigned for "; Discipline.ID; disciplineTags.Count
        
        Set discWs = Discipline.CreateDisciplineOutputSheet(ThisWorkbook.Sheets)
        disciplineTags.OutputTagListings discWs.Range("A1")
    Next Discipline
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
    
'Read in FailureCodes
    Set ws = wb.Worksheets(FailureCodesSheetName)
    Set tbl = ws.ListObjects(FailureCodesTableName)
    Set FailureCodes = New clsFailureCodes
    FailureCodes.LoadTable tbl
    
'TODO Read in Ex list - optional, if not added to tag.

End Sub
