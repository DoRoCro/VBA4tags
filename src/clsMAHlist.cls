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

Sub LoadTableProcess(table As ListObject)
    Dim MAHDefault As clsMAHDefault
    Dim row As ListRow  'for use in for Each

    For Each row In table.ListRows
        'read an entry
        Set MAHDefault = New clsMAHDefault
        'Debug.Print row.index, row.Range(1, 1)
        With MAHDefault
            .ID = row.Range(1, 1)
        End With
        'add it to collection
        Me.Add MAHDefault
    Next
End Sub
