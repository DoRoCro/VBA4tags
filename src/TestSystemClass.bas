Attribute VB_Name = "TestSystemClass"
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
Public Sub TestSystemHasID()
    On Error GoTo TestFail
    
    'Arrange:
    Dim tstSystem As clsSystem
    
    'Act:
    Set tstSystem = New clsSystem
    tstSystem.SystemID = 24        'Flash_Gas_Compression
    'Assert:
    Assert.istrue (tstSystem.SystemID = 24)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub TestSystemHasDescriptionAndUtilityTest()
    On Error GoTo TestFail
    
    'Arrange:
    Dim tstSystem As clsSystem
    Set tstSystem = New clsSystem

    'Act:
    With tstSystem
        tstSystem.Description = "Flash Gas Compression"
        'tstSystem.FluidType = "Hydrocarbons"
        tstSystem.isUtility = False
        tstSystem.SystemNumber = "24"
    End With
    'Assert:
    Assert.istrue (tstSystem.Description = "Flash Gas Compression")
    'Assert.isTrue (tstSystem.FluidType = "Hydrocarbons")
    Assert.istrue (tstSystem.isUtility = False)
    Assert.istrue (tstSystem.SystemNumber = "24")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

''@TestMethod
'Public Sub TestLoadSystems()
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim Systems As New Collection
'    Dim System As New clsSystem
'    'Act:
'
'
'    'Assert:
'    Assert.isTrue (Systems.Count = 1)
'
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'End Sub

'@TestMethod
Public Sub TestLoadSystemsFromTable()
    On Error GoTo TestFail
    
    'Arrange:
    Dim Systems As clsSystems
    Set Systems = New clsSystems
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("TestSystems").ListObjects("TestSystemsList")
    
    'Act:
    Systems.LoadTable tbl
    Debug.Print Systems.Item(1).SystemID
    'Assert:
    Assert.istrue (Systems.Item(1).SystemID = "E-SYSTEM-00")
    Assert.istrue (Systems.Item(100).SystemID = "E-SYSTEM-99")
    Assert.istrue (Systems.Item(101).SystemID = "R-SYSTEM-00")
    Assert.istrue (Systems.Item(200).SystemID = "R-SYSTEM-99")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


