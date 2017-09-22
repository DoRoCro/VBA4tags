VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFailureCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CriticalityAssignment")
Private pContent As New Collection

Sub LoadTable(table As ListObject)
    Dim FailureCode As clsFailureCode
    Dim row As ListRow  'for use in for Each

    For Each row In table.ListRows
        If Not WorksheetFunction.IsErr(rowCell(row, "Number found in ASSET-C WND")) Then
            'read an entry
            Set FailureCode = New clsFailureCode
            'Debug.Print row.index, row.Range(1, 1)
            With FailureCode
                .ID = row.Range(1, 1)
                .Description = row.Range(1, 2)
            End With
            'add it to collection
            Me.Add FailureCode
        End If
    Next
End Sub

Sub Add(newFailureCode As clsFailureCode)
    pContent.Add newFailureCode
End Sub

Function Item(index As LongPtr) As clsFailureCode
    Set Item = pContent.Item(index)
End Function

Function Count() As LongPtr
    Count = pContent.Count
End Function

Public Function All() As Collection
    Set All = pContent
End Function

'Public Sub createOutputSheetsByFailureCode() ' TODO add check and delete/rename if existing FailureCode worksheet
'    Dim FailureCode As clsFailureCode
'    For Each FailureCode In pContent
'      Set FailureCode.OutputSheet = FailureCode.CreateFailureCodeOutputSheet(ThisWorkbook.Sheets)
'    Next
'End Sub



