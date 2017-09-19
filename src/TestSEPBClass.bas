Attribute VB_Name = "TestSEPBClass"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod
Public Sub TestSEPBInput()
    On Error GoTo TestFail
    
    'Arrange:
    Dim SEPB As New clsSEPB
    Set SEPB = New clsSEPB
    
    'Act:
    SEPB.Impact(Safety) = "A"
    SEPB.Likelihood(Safety) = 1
    'Assert:
    Assert.AreEqual SEPB.Risk(Safety), "A1"
    'Assert.Inconclusive

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

