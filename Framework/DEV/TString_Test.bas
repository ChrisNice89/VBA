Attribute VB_Name = "TString_Test"
Option Compare Database
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider
Private SUT As TString

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Uncategorized")
Private Sub TStringInitialize()
    On Error GoTo TestFail
    
    'Arrange:
    
    'Act:
    Set SUT = TString("TestValue")
    
    'Assert:
    Assert.AreNotEqual VarPtr(SUT), VarPtr(TString)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TStringInitialize_Instances()
    On Error GoTo TestFail
    
    'Arrange:
    Dim s1 As TString
    Dim s2 As TString
    Dim i As Long
    'Act:
    For i = 1 To 10000
        Set s1 = TString("TestValue" & i)
    Next
    
    Set s1 = TString("TestValue")
    Set s2 = TString("TestValue2")
    'Assert:
    Assert.AreNotEqual VarPtr(s1), VarPtr(s2)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TStringValue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim result As String
    'Act:
    Set SUT = TString("TestValue")
    result = SUT.Value
    'Assert:
    Assert.AreEqual "TestValue", result

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Uncategorized")
Private Sub TStringValue2()
    On Error GoTo TestFail
    
    'Arrange:
    Dim s1 As TString
    Dim s2 As TString
    
    'Act:
    Set s1 = TString("TestValue")
    Set s2 = TString("TestValue2")
    'Assert:
    Assert.AreNotEqual s1.Value, s2.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Uncategorized")
Private Sub TStringValue2_Equal()
    On Error GoTo TestFail
    
    'Arrange:
    Dim s1 As TString
    Dim s2 As TString
    
    'Act:
    Set s1 = TString("TestValue")
    Set s2 = TString("TestValue")
    'Assert:
    Assert.AreEqual s1.Value, s2.Value

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
