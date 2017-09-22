VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFailureCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CriticalityAssignment")
Option Explicit

Private pID As String
Private pDescription As String
Private pOutputSheet As Worksheet

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

Public Property Get OutputSheet() As Worksheet
    Set OutputSheet = pOutputSheet
End Property
Public Property Set OutputSheet(Value As Worksheet)
    Set pOutputSheet = Value
End Property
'Public Function CreateFailureCodeOutputSheet(wbsheets As Sheets) As Worksheet
'    'TODO add guards if FailureCode worksheet already exists
'    Dim newWs As Worksheet
'    If Me.ID = "" Then Me.ID = "BLANK"
'    If Me.ID = "N/A" Then Me.ID = "N_A"
'    Set newWs = wbsheets.Add(After:=Sheets(Sheets.Count))
'    newWs.Name = Me.ID
'    Set Me.OutputSheet = wbsheets(Me.ID)
'    Set CreateFailureCodeOutputSheet = newWs
'End Function

