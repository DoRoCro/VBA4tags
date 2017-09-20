VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("CriticalityAssignment")

Private pSystemID As String
Private pDescription As String
Private pFluidType As String
Private pSystemNumber As String

Private pIsUtility As Boolean

Public Property Get SystemID() As String
    SystemID = pSystemID
End Property
Public Property Let SystemID(Value As String)
    pSystemID = Value
End Property

Public Property Get Description() As String
    Description = pDescription
End Property
Public Property Let Description(Value As String)
    pDescription = Value
End Property
Public Property Get SystemNumber() As String
    SystemNumber = pSystemNumber
End Property
Public Property Let SystemNumber(Value As String)
    pSystemNumber = Value
End Property

Public Property Get FluidType() As String
    FluidType = pFluidType
End Property
Public Property Let FluidType(Value As String)
    Set pFluidType = Value
End Property

Public Property Get isUtility() As String
    isUtility = pIsUtility
End Property
Public Property Let isUtility(Value As String)
    pIsUtility = Value
End Property

