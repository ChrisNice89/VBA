Attribute VB_Name = "Modul1"
'@Folder "Entwicklung"

Option Explicit
Sub Compare()


    Dim k As TString
    Dim v As TString
    Dim i As Long
    Dim A(1 To 1000) As IObject
    Dim B(1 To 1000) As IObject
    
    For i = 1 To 1000
        Set A(i) = TString.Build("A")
        Set B(i) = TString.Build("B")
    Next
    
    Dim t As New Timer
    t.StartCounter
    For i = 1 To 1000
        If Not A(i).CompareTo(B(i)) = IsLower Then
            Debug.Print "error"
        End If
    Next
    Debug.Print t.TimeElapsed



End Sub

Sub Equals()

    Dim k As TString
    Dim v As TString
    Dim i As Long
    Dim A(1 To 1000) As IObject
    Dim B(1 To 1000) As IObject
    
    For i = 1 To 1000
        Set A(i) = TString.Build("Key" & i)
        Set B(i) = TString.Build("xey" & i)
    Next
    
    Dim t As New Timer
    t.StartCounter
    For i = 1 To 1000
        Call A(i).Equals(B(i))
    Next
    Debug.Print t.TimeElapsed

End Sub

Sub test()

    Dim k As TString
    Dim v As TString
    Dim i As Long
    
    Dim map As TMap
    Set map = TMap.Build(TString, TString)
    Dim t As New Timer
    t.StartCounter
    For i = 1 To 10000
        Set k = TString.Build("Key" & i)
        Set v = TString.Build("Value" & i)
        Call map.Add(k, v)
    Next
    Debug.Print t.TimeElapsed
    Debug.Print Skynet.IObject(map).ToString
End Sub


Sub ttt()


    Dim k As TString
    Dim v As TString
    Dim i As Long
    Dim t As New Timer
    
    t.StartCounter
    For i = 1 To 10000
        Set v = TString.Build("Value" & i)
    Next
    Debug.Print t.TimeElapsed

End Sub

Public Sub Is64BitExcel()


#If VBA7 And Win64 Then
    Debug.Print True
#Else
    Debug.Print False
#End If
End Sub
