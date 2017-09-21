VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSystems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("CriticalityAssignment")

Private pContent As New Collection

Sub LoadTable(table As ListObject)
    Dim system As clsSystem
    Dim row As ListRow  'for use in for Each


    For Each row In table.ListRows
        'read an entry
        Set system = New clsSystem
        'Debug.Print row.index, row.Range(1, 1)
        With system
            .SystemID = row.Range(1, 1)
            .Description = row.Range(1, 2)
            .SystemNumber = row.Range(1, 3) 'as string
            If row.Range(1, 5) = "N" Then
                .isUtility = False
            Else
                .isUtility = True
            End If
        End With
        'add it to collection
        Me.Add system
        
    Next
End Sub

Sub Add(newSystem As clsSystem)
    pContent.Add newSystem, newSystem.SystemID
End Sub

Function Item(index As LongPtr) As clsSystem
    Set Item = pContent.Item(index)
End Function

Function FindByNumber(SystemNumber As String) As clsSystem
    Set FindByNumber = pContent.Item("E-SYSTEM-" & SystemNumber)
End Function

Function Count() As LongPtr
    Count = pContent.Count
End Function

Function Contains(SystemID As String) As Boolean
    Dim system As clsSystem
    On Error GoTo err
    Contains = True
    Set system = pContent.Item("E-SYSTEM-" & SystemID)
    Exit Function
    
err:
    Contains = False
End Function
