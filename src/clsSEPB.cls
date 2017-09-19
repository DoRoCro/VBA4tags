VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSEPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Private pImpacts(4) As String * 1
Private pLikelihoods(4) As String * 1
Public Enum SEPB
    Safety
    Environment
    Production
    Business
End Enum
    
Public Property Get Impact(Category As SEPB) As String
    Impact = pImpacts(Category)
End Property
Public Property Let Impact(Category As SEPB, Value As String)
    pImpacts(Category) = Value
End Property
    
Public Property Get Likelihood(Category As SEPB) As String
    Likelihood = pLikelihoods(Category)
End Property
Public Property Let Likelihood(Category As SEPB, Value As String)
    pLikelihoods(Category) = Value
End Property

Public Property Get Risk(Category As SEPB) As String
    Risk = pImpacts(Category) & pLikelihoods(Category)
End Property

