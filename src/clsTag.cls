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
Private pJustification As String


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

Public Property Get isSIL() As Boolean
    isSIL = pisSIL
End Property
Public Property Let isSIL(Value As Boolean)
    pisSIL = Value
End Property

Public Property Get isSIS() As Boolean
    isSIS = pisSIS
End Property
Public Property Let isSIS(Value As Boolean)
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

Public Property Get Justification() As String
    Justification = pJustification
End Property
Public Property Let Justification(Value As String)
    pJustification = Value
End Property



Private Sub class_initialize()
  Set pRisks = New clsSEPB
End Sub

Public Sub WriteToWorksheet(StartCell As Range)
    With Me
        StartCell.Formula = .ID
        StartCell.Offset(0, 1).Formula = .Description
        StartCell.Offset(0, 2).Formula = .FailureCode
        StartCell.Offset(0, 3).Formula = .Status
        StartCell.Offset(0, 4).Formula = .UMCLplus
        StartCell.Offset(0, 5).Formula = .ISOcode
        StartCell.Offset(0, 6).Formula = .isSIL
        StartCell.Offset(0, 7).Formula = .isSIS
        StartCell.Offset(0, 8).Formula = .Discipline
        StartCell.Offset(0, 9).Formula = .SystemID
        StartCell.Offset(0, 10).Formula = .Criticality
        StartCell.Offset(0, 11).Formula = .Risks(Safety)
        StartCell.Offset(0, 12).Formula = .Risks(Environment)
        StartCell.Offset(0, 13).Formula = .Risks(Production)
        StartCell.Offset(0, 14).Formula = .Risks(Business)
        StartCell.Offset(0, 15).Formula = .Justification
        StartCell.Offset(0, 16).Formula = .Criticality
    End With
End Sub
Public Sub FillTagHeaders(StartCell As Range)
    StartCell.Formula = "Functional Location"
    StartCell.Offset(0, 1).Formula = "Functional Loccation Description"
    StartCell.Offset(0, 2).Formula = "FailureCode"
    StartCell.Offset(0, 3).Formula = "Status"
    StartCell.Offset(0, 4).Formula = "UMCLplus"
    StartCell.Offset(0, 5).Formula = "ISOcode"
    StartCell.Offset(0, 6).Formula = "isSIL"
    StartCell.Offset(0, 7).Formula = "isSIS"
    StartCell.Offset(0, 8).Formula = "discipline"
    StartCell.Offset(0, 9).Formula = "SystemID"
    StartCell.Offset(0, 10).Formula = "Initial Criticality"
    StartCell.Offset(0, 11).Formula = "Risks(Safety)"
    StartCell.Offset(0, 12).Formula = "Risks(Environment)"
    StartCell.Offset(0, 13).Formula = "Risks(Production)"
    StartCell.Offset(0, 14).Formula = "Risks(Business)"
    StartCell.Offset(0, 15).Formula = "Justification"
    StartCell.Offset(0, 14).Formula = "Risks(Business)"
    
    StartCell.Offset(0, 16).Formula = "Criticality to allocate"
    StartCell.Offset(0, 17).Formula = "Reason for change"
    
End Sub

Sub SetDefaultCriticalityByFailureCode()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim MAHCell As Range
    Dim resetComponent As String
    Dim resetComment As String
    Set wb = Workbooks(CriticalityWbName)
    Select Case Me.FailureCode
        Case "SOFT", "LOOP", vbNullString      'ignore these failure codes
            Me.Criticality = "F"
        Case Else
            Set ws = wb.Worksheets(Me.FailureCode)
            Set MAHCell = ws.Range("I17")
            Me.Criticality = ws.Range("K1")
            Me.RiskImpact(Safety) = ws.Range("B9")
            Me.RiskImpact(Environment) = ws.Range("B10")
            Me.RiskImpact(Production) = ws.Range("B11")
            Me.RiskImpact(Business) = ws.Range("B12")
            Me.RiskLikelihood(Safety) = ws.Range("C9")
            Me.RiskLikelihood(Environment) = ws.Range("C10")
            Me.RiskLikelihood(Production) = ws.Range("C11")
            Me.RiskLikelihood(Business) = ws.Range("C12")
            Me.Justification = "MAH Barrier: " & ws.Range("H17").Text & "; " & _
                                "MAH comment: " & ws.Range("I19").Text & "; " & _
                                "; Regulatory override = " & ws.Range("I37").Text
            If Me.isSIL Or Me.isSIS Then
                Me.Justification = Me.Justification & "; In SIL/SIS Loop, therefore criticality = A"
                Me.Criticality = "A"
            End If
    End Select
End Sub

