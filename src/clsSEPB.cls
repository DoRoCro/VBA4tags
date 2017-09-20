VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSEPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("CriticalityAssignment")
Private pImpacts(4) As String * 1
Private pLikelihoods(4) As String * 1
Public Enum SEPB
    Safety
    Environment
    Production
    Business
End Enum
    
Public Property Get Impact(category As SEPB) As String
    Impact = pImpacts(category)
End Property
Public Property Let Impact(category As SEPB, Value As String)
    pImpacts(category) = Value
End Property
    
Public Property Get Likelihood(category As SEPB) As String
    Likelihood = pLikelihoods(category)
End Property
Public Property Let Likelihood(category As SEPB, Value As String)
    pLikelihoods(category) = Value
End Property

Public Property Get Risk(category As SEPB) As String
    Risk = pImpacts(category) & pLikelihoods(category)
End Property

