VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDiscipline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Private pID As String
Private pDescription As String


Public Property Get ID() As String
    ID = pID
End Property
Public Property Let ID(Value As String)
    pID = Value
End Property
