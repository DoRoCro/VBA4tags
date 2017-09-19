VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Class clsTag
'Attributes

Private pID As String
Private pDescription As String
Private pFailureCode As String
Private pStatus As String
Private pUMCLplus As String
Private pISOcode As String
Private pisSIL As Boolean
Private pisSIS As Boolean
Private pDiscipline As String
Private pSystemID As String
Private pCriticality As String
Private pRisks As clsSEPB


Public Property Get ID() As String
    ID = pID
End Property
Public Property Let ID(Value As String)
    pID = Value
End Property

Public Property Get Description() As String
    Description = pDescription
End Property
Public Property Let Description(Value As String)
    pDescription = Value
End Property

Public Property Get FailureCode() As String
    FailureCode = pFailureCode
End Property
Public Property Let FailureCode(Value As String)
    pFailureCode = Value
End Property

Public Property Get Status() As String
    Status = pStatus
End Property
Public Property Let Status(Value As String)
    pStatus = Value
End Property

Public Property Get UMCLplus() As String
    UMCLplus = pUMCLplus
End Property
Public Property Let UMCLplus(Value As String)
    pUMCLplus = Value
End Property

Public Property Get ISOcode() As String
    ISOcode = pISOcode
End Property
Public Property Let ISOcode(Value As String)
    pISOcode = Value
End Property

Public Property Get Discipline() As String
    Discipline = pDiscipline
End Property
Public Property Let Discipline(Value As String)
    pDiscipline = Value
End Property

Public Property Get SystemID() As String
    SystemID = pSystemID
End Property
Public Property Let SystemID(Value As String)
    pSystemID = Value
End Property

Public Property Get isSIL() As String
    isSIL = pisSIL
End Property
Public Property Let isSIL(Value As String)
    pisSIL = Value
End Property

Public Property Get isSIS() As String
    isSIS = pisSIS
End Property
Public Property Let isSIS(Value As String)
    pisSIS = Value
End Property


Public Property Get Criticality() As String
    Criticality = pCriticality
End Property
Public Property Let Criticality(Value As String)
    pCriticality = Value
End Property

Public Property Get Risks(category As SEPB) As String
    Risks = pRisks.Risk(category)
End Property
Public Property Let RiskImpact(category As SEPB, Value As String)
    pRisks.Impact(category) = Value
End Property
Public Property Let RiskLikelihood(category As SEPB, Value As String)
    pRisks.Likelihood(category) = Value
End Property

Private Sub Class_Initialize()
  Set pRisks = New clsSEPB
End Sub

Public Sub WriteToWorksheet(StartCell As Range)
    With Me
        StartCell.Select
        StartCell.Formula = .ID
'        StartCell.
'pFailureCode As String
'pStatus As String
'pUMCLplus As String
'pISOcode As String
'pisSIL As Boolean
'pisSIS As Boolean
'pDiscipline As String
'pSystemID As String
'pCriticality As String
'pRisks As clsSEPB
    
    End With
    
End Sub
