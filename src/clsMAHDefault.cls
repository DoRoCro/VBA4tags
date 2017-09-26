VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMAHDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CriticalityAssignment")
Option Explicit

Private pID As String
Private pFamilyID As String
Private pFamily As String
Private pComponent As String
Private pComment As String
Private pTypCriticality As String * 1
'initially going to have 2 barriers instances, one for process, one for utility systems

Public Property Get ID() As String
    ID = pID
End Property
Public Property Let ID(Value As String)
    pID = Value
End Property

Public Property Get FamilyID() As String
    FamilyID = pFamilyID
End Property
Public Property Let FamilyID(Value As String)
    pFamilyID = Value
End Property

Public Property Get Family() As String
    Family = pFamily
End Property
Public Property Let Family(Value As String)
    pFamily = Value
End Property

Public Property Get Component() As String
    Component = pComponent
End Property
Public Property Let Component(Value As String)
    pComponent = Value
End Property

Public Property Get Comment() As String
    Comment = pComment
End Property
Public Property Let Comment(Value As String)
    pComment = Value
End Property

Public Property Get TypCriticality() As String
    TypCriticality = pTypCriticality
End Property
Public Property Let TypCriticality(Value As String)
    pTypCriticality = Value
End Property

