VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")

Private pSystemID As Integer
Private pDescription As String
Private pFluidType As String

Public Property Get SystemID() As Integer
    SystemID = pSystemID
End Property
Public Property Let SystemID(Value As Integer)
    pSystemID = Value
End Property

Public Property Get Description() As String
    Description = pDescription
End Property
Public Property Let Description(Value As String)
    pDescription = Value
End Property

Public Property Get FluidType() As String
    FluidType = pFluidType
End Property
Public Property Let FluidType(Value As String)
    pFluidType = Value
End Property

