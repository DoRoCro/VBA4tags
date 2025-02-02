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
Private Disciplines As clsDisciplines
Private Systems As clsSystems
Private MAHprocess As clsMAHlist
Private MAHutility As clsMAHlist
Private FailureCodes As clsFailureCodes


Sub AssignCriticalities()
    Dim tag As clsTag
    Dim disciplineTags As clsTags
    Dim discWs As Worksheet
    Dim Discipline As clsDiscipline
    
    CriticalityWbName = "WND Criticality Template.xlsx"
    Application.DisplayStatusBar = True
    Application.StatusBar = "Loading from tables..."
    Set tags = New clsTags
    Set Systems = New clsSystems
    Set Disciplines = New clsDisciplines
    Set FailureCodes = New clsFailureCodes
    Call LoadTables(tags, Systems, Disciplines)
    Debug.Print "Tags count ="; tags.Count
    Debug.Print "Systems count ="; Systems.Count
    Debug.Print "Disciplines count ="; Disciplines.Count
    Debug.Print "Systems count ="; Systems.Count
    Debug.Print "Failurecodes count ="; FailureCodes.Count
    Debug.Print "MAHprocess count ="; MAHprocess.Count
    Debug.Print "MAHutility count ="; MAHutility.Count
    
    'filter out unwanted tags by STATUS code
    Application.StatusBar = "Filtering unused tags..."
    Set tags = tags.RemoveStatus("DEL")
    Set tags = tags.RemoveStatus("SOFT")
    Set tags = tags.RemoveStatus(vbNullString)
    Set tags = tags.RemoveStatus("DRAFT")
    
    Set disciplineTags = New clsTags

    For Each Discipline In Disciplines.All
        Application.StatusBar = "Processing " & Discipline.ID & " tags..."

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
        
        Application.StatusBar = "Writing out " & Discipline.ID & " tags..."
       
        Set discWs = Discipline.CreateDisciplineOutputSheet(ThisWorkbook.Sheets)
        discWs.Range("J1").Formula = "Count A ="
        discWs.Range("J2").Formula = "Count B ="
        discWs.Range("J3").Formula = "Count C ="
        discWs.Range("K1").Formula = "=COUNTIF(K6:K12000,""A"")"
        discWs.Range("K2").Formula = "=COUNTIF(K6:K12000,""B"")"
        discWs.Range("K3").Formula = "=COUNTIF(K6:K12000,""C"")"
        
        
        disciplineTags.OutputTagListings discWs.Range("A5")
    Next Discipline
    Application.StatusBar = ""
End Sub


Sub LoadTables(tags As clsTags, Systems As clsSystems, Disciplines As clsDisciplines)
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
    Set Disciplines = New clsDisciplines
    Disciplines.LoadTable tbl
    Debug.Print "finished loading disciplines, count ="; Disciplines.Count

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
Sub MoveSheetsToDiscplineWorkbook()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsname As String
    Set wb = ActiveWorkbook
    For Each ws In wb.Worksheets
        ws.Select
        ws.Copy
        wsname = ws.Name
        ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\Douglas.Crooke\OneDrive - add energy group\Discipline - " & wsname & ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    wb.Activate
    Next ws
End Sub


