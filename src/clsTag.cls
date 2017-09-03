VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class clsTag
'Attributes

Private ptagID As String
Private ptagDescription As String

Public Property Get TagID() As String
    TagID = ptagID
End Property
Public Property Let TagID(Value As String)
    ptagID = Value
End Property

Public Property Get TagDescription() As String
    TagDescription = ptagDescription
End Property
Public Property Let TagDescription(Value As String)
    ptagDescription = Value
End Property

