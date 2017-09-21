VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMAHlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("CriticalityAssignment")

Private pContent As New Collection

Public Sub Add(MAH As clsMAHDefault)
    pContent.Add MAH, MAH.ID
End Sub

Public Function Item(index As LongPtr) As clsMAHDefault
    Set Item = pContent.Item(index)
End Function
Public Function FindByID(ID As String) As clsMAHDefault
    Set FindByID = pContent.Item(ID)
End Function



Public Function Count() As LongPtr
    Count = pContent.Count
End Function

Public Sub LoadTableProcess(table As ListObject)
    Dim MAHDefault As clsMAHDefault
    Dim row As ListRow  'for use in for Each

    For Each row In table.ListRows
        'read an entry
        Set MAHDefault = New clsMAHDefault
        'Debug.Print row.index, row.Range(1, 1)
        With MAHDefault
            .ID = row.Range(1, 1).Text
            .Family = row.Range(1, 4).Text          'to come with #N/A entries
            .Component = row.Range(1, 5).Text
            .TypCriticality = row.Range(1, 6).Text
            .Comment = row.Range(1, 7).Text
        End With
        'add it to collection
        Me.Add MAHDefault
    Next
End Sub

Public Sub LoadTableUtility(table As ListObject)
    Dim MAHDefault As clsMAHDefault
    Dim row As ListRow  'for use in for Each

    For Each row In table.ListRows
        'read an entry
        Set MAHDefault = New clsMAHDefault
        'Debug.Print row.index, row.Range(1, 1)
        With MAHDefault
            .ID = row.Range(1, 1).Text
            .Family = row.Range(1, 9).Text          'to come with #N/A entries
            .Component = row.Range(1, 10).Text
            .TypCriticality = row.Range(1, 11).Text
            .Comment = row.Range(1, 12).Text
        End With
        'add it to collection
        Me.Add MAHDefault
    Next
End Sub
