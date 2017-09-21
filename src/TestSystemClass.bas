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
    Assert.isTrue (tstSystem.SystemID = 24)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
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
    Assert.isTrue (tstSystem.Description = "Flash Gas Compression")
    'Assert.isTrue (tstSystem.FluidType = "Hydrocarbons")
    Assert.isTrue (tstSystem.isUtility = False)
    Assert.isTrue (tstSystem.SystemNumber = "24")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
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
    Assert.isTrue (Systems.Item(1).SystemID = "E-SYSTEM-00")
    Assert.isTrue (Systems.Item(100).SystemID = "E-SYSTEM-99")
    Assert.isTrue (Systems.Item(101).SystemID = "R-SYSTEM-00")
    Assert.isTrue (Systems.Item(200).SystemID = "R-SYSTEM-99")
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub

'@TestMethod
Public Sub TestSystemsContains() 'TODO Rename test
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
    Assert.isFalse Systems.Contains("N/A")
    Assert.isTrue Systems.Contains("24")

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & err.Number & " - " & err.Description
End Sub


