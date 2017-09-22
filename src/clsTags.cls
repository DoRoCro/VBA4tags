VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("CriticalityAssignment")

Private pContent As New Collection

Sub LoadTable(table As ListObject)
    Dim tag As clsTag
    Dim row As ListRow  'for use in for Each


    For Each row In table.ListRows
        'read an entry
        Set tag = New clsTag
        'Debug.Print row.index, row.Range(1, 1)
        With tag
            .ID = row.Range(1, 1)
            .Description = row.Range(1, 2)
            .FailureCode = row.Range(1, 3) 'as string
            .Status = row.Range(1, 4)
            .UMCLplus = row.Range(1, 5)
            .ISOcode = row.Range(1, 6)
            If row.Range(1, 7) = "Y" Or row.Range(1, 8) = "Y" Then
                .isSIL = True
            Else
                .isSIL = False
            End If
            If row.Range(1, 9) = "X" Then
                .isSIS = True
            Else
                .isSIS = False
            End If
            .Discipline = row.Range(1, 10)
            .SystemID = row.Range(1, 11)
        End With
        'add it to collection
        Me.Add tag
        
    Next
End Sub
Sub LoadArray(table As Variant)
    Dim tag As clsTag
    Dim row As LongPtr  'for use in for Each

    For row = LBound(table) To UBound(table)
        'read an entry
        Set tag = New clsTag
        'Debug.Print row.index, row.Range(1, 1)
        With tag
            .ID = table(row, 1)
            .Description = table(row, 2)
            .FailureCode = table(row, 3) 'as string
            .Status = table(row, 4)
            .UMCLplus = table(row, 5)
            .ISOcode = table(row, 6)
            If table(row, 7) = "Y" Or table(row, 8) = "Y" Then
                .isSIL = True
            Else
                .isSIL = False
            End If
            If table(row, 9) = "X" Then
                .isSIS = True
            Else
                .isSIS = False
            End If
            .Discipline = table(row, 10)
            .SystemID = table(row, 11)
        End With
        'add it to collection
        Me.Add tag
        
    Next
End Sub
Sub Add(newTag As clsTag)
    pContent.Add newTag
End Sub

Function Item(index As LongPtr) As clsTag
    Set Item = pContent.Item(index)
End Function

Function Count() As LongPtr
    Count = pContent.Count
End Function

Sub OutputTagListings(StartCell As Range)
    Dim tag As clsTag
    Dim counter As Long
    Dim outputArray() As Variant
    ReDim outputArray(1 To Me.Count, 12)
    counter = 1
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Me.Item(1).FillTagHeaders StartCell      'dummy use of a tag use to use class method
    For Each tag In pContent
        tag.WriteToWorksheet StartCell.Offset(counter, 0)
        counter = counter + 1
    Next
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


Public Function All() As Collection
    Set All = pContent
End Function

Public Function byDiscipline(Discipline As clsDiscipline) As clsTags
    Dim tag As clsTag
    Dim tagsByDiscipline As clsTags
    Set tagsByDiscipline = New clsTags
    For Each tag In Me.All
        If tag.Discipline = Discipline.ID Then tagsByDiscipline.Add tag
    Next tag
    Set byDiscipline = tagsByDiscipline
End Function

Public Function RemoveStatus(StatusCode As String) As clsTags
    Dim tag As clsTag
    Dim tagsWithoutStatus As clsTags
    Set tagsWithoutStatus = New clsTags
    For Each tag In Me.All
        If tag.Status <> StatusCode Then tagsWithoutStatus.Add tag
    Next tag
    Set RemoveStatus = tagsWithoutStatus
End Function
Public Function ProcessTags(Systems As clsSystems) As clsTags
    Dim tag As clsTag
    Dim TagsProcess As clsTags
    Set TagsProcess = New clsTags
    For Each tag In Me.All
        If Systems.Contains(tag.SystemID) Then
            If Not Systems.FindByNumber(tag.SystemID).isUtility Then TagsProcess.Add tag
        End If
    Next tag
    Set ProcessTags = TagsProcess
End Function
Public Function UtilityTags(Systems As clsSystems) As clsTags
    Dim tag As clsTag
    Dim TagsUtility As clsTags
    Set TagsUtility = New clsTags
    For Each tag In Me.All
        If Systems.Contains(tag.SystemID) Then
            If Systems.FindByNumber(tag.SystemID).isUtility Then TagsUtility.Add tag
        End If
    Next tag
    Set UtilityTags = TagsUtility
End Function
Public Function NoSystemTags(Systems As clsSystems) As clsTags
    Dim tag As clsTag
    Dim TagsNoSystem As New clsTags
    For Each tag In Me.All
        If Not Systems.Contains(tag.SystemID) Then TagsNoSystem.Add tag
    Next tag
    Set NoSystemTags = TagsNoSystem
End Function

Public Sub AssignDefaultCriticalities()
    Dim tag As clsTag
    For Each tag In Me.All
        tag.SetDefaultCriticalityByFailureCode
    Next tag
End Sub
