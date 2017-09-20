Attribute VB_Name = "TestMAHBarrierDefaults"
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
Public Sub TestMAHbasics()
    On Error GoTo TestFail
    
    'Arrange:
    Dim MAH As clsMAHDefaults
    Set MAH = New clsMAHDefaults
    'Act:
    MAH.ID = "FA_CFBC"
    MAH.Component = "Some text here"
    MAH.Family = "Family"
    MAH.Comment = "Comment"
    MAH.TypCriticality = "D"

    'Assert:
    Assert.AreEqual MAH.ID, "FA_CFBC"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

