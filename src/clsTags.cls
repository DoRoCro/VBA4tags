VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")

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

Sub Add(newTag As clsTag)
    pContent.Add newTag
End Sub

Function Item(index As LongPtr) As clsTag
    Set Item = pContent.Item(index)
End Function

Function Count() As LongPtr
    Set Count = pContent.Count
End Function
