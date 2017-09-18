VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSystems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Tests")

Private pContent As New Collection

Sub Load(table As ListObject)
    Dim System As clsSystem
    Set System = New clsSystem
    'read an entry
    System.SystemID = 50
    'add it to collection
    pContent.Add System
End Sub

