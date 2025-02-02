VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDisciplines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("CriticalityAssignment")
Private pContent As New Collection

Sub LoadTable(table As ListObject)
    Dim Discipline As clsDiscipline
    Dim row As ListRow  'for use in for Each

    For Each row In table.ListRows
        'read an entry
        Set Discipline = New clsDiscipline
        'Debug.Print row.index, row.Range(1, 1)
        With Discipline
            .ID = row.Range(1, 1)
        End With
        'add it to collection
        Me.Add Discipline
    Next
End Sub

Sub Add(newDiscipline As clsDiscipline)
    pContent.Add newDiscipline
End Sub

Function Item(index As LongPtr) As clsDiscipline
    Set Item = pContent.Item(index)
End Function

Function Count() As LongPtr
    Count = pContent.Count
End Function

Public Function All() As Collection
    Set All = pContent
End Function

Public Sub createOutputSheetsByDiscipline() ' TODO add check and delete/rename if existing discipline worksheet
    Dim Discipline As clsDiscipline
    For Each Discipline In pContent
      Set Discipline.OutputSheet = Discipline.CreateDisciplineOutputSheet(ThisWorkbook.Sheets)
    Next
End Sub


