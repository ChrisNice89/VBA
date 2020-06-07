Attribute VB_Name = "TMap_Test"
Option Compare Database
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider
Private SUT As TMap

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
Private Sub TMapInitialize()
    On Error GoTo TestFail
    
    'Arrange:
    
    'Act:
    Set SUT = TMap.Build(TString, TString)
    
    'Assert:
    Assert.AreNotEqual VarPtr(SUT), VarPtr(TMap)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TMap_Two_Instances()
    On Error GoTo TestFail
    
    'Arrange:
    Dim sut1 As TMap
    Dim sut2 As TMap
    
    'Act:
    Set sut1 = TMap.Build(TString, TString)
    Set sut2 = TMap.Build(TString, TString)
    'Assert:
    Assert.AreNotEqual VarPtr(sut1), VarPtr(sut2)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TMapOverall_2()
    On Error GoTo TestFail
    
    'Arrange:
    Dim i As Long
    Dim k As TString
    Dim v As TString
    
    'Act:
   
    Set SUT = TMap.Build(TString, TString, 0, 0.7)
    
    For i = 1 To 10000
        Set k = TString("Key" & i)
        Set v = TString("Value" & i)
        Call SUT.Add(k, v)
    Next
    
    Assert.Succeed
   
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TMapValue_AddThree_TSTring()
    On Error GoTo TestFail
    
    'Arrange:
    Dim s1 As TString
    Dim s2 As TString
    Dim s3 As TString
    
    'Act:
    Set s1 = TString("TestKey")
    Set s2 = TString("TestKey2")
    Set s3 = TString("TestKey3")
    
    Set SUT = TMap.Build(TString, TString)
    
    Call SUT.Add(s1, TString("TestValue"))
    SUT.Item(s2) = TString("TestValue2")
    Set SUT.Item(s3) = TString("TestValue3")
    
    'Assert:
    Assert.IsTrue SUT.Contains(s1)
    Assert.IsTrue SUT.Contains(s2)
    Assert.IsTrue SUT.Contains(s3)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TMapOverall()
    On Error GoTo TestFail
    
    'Arrange:
    Dim i As Long
    Dim k As TString
    Dim v As TString
    
    'Act:
   
    Set SUT = TMap.Build(TString, TString)
    
    For i = 1 To 500
        Set k = TString("Key" & i)
        Set v = TString("Value" & i)
        Call SUT.Add(k, v)
    Next
    Assert.IsTrue (SUT.Count = 500)
    
    For i = 501 To 1000
        Set k = TString("Key" & i)
        Set v = TString("Value" & i)
        Set SUT.Item(k) = v
    Next
    
    Assert.IsTrue (SUT.Count = 1000)
    
    For i = 1 To 1000
        Set k = TString("Key" & i)
        If SUT.Contains(k) Then
            Set v = SUT.LastCheck
            Assert.AreEqual "Value" & i, v.Value
        Else
            Assert.Fail ("SUT.Contains(TString(Key & i))")
        End If
        
        Set v = SUT.Item(k)
        Assert.AreEqual "Value" & i, v.Value
    Next
    
    
    
    Assert.IsFalse SUT.Contains(TString("Key" & i))
    Set v = SUT.Item(TString("Key" & i))
    Assert.IsNothing v
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
