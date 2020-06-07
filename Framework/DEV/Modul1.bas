Attribute VB_Name = "Modul1"
'@Folder "Entwicklung"

Option Explicit

Sub ParameterTest()

Dim cmd As SqlCommand
Set cmd = SqlCommand.Build("SomeSql", SqlConnection)

Call cmd.CreateParameter(TString.Build("Christian"), "Name").AddValue(TString.Build("Christoph"))
Debug.Print cmd.Parameter("Name").Current.Value

Debug.Print cmd.Parameter("Name").UseValue(2).Object.ToString

Dim Christian As IObject
Dim Christoph As IObject

Set Christian = cmd.Parameter("Name").Value(1)
Set Christoph = cmd.Parameter("Name").Value(2)

Debug.Print Christoph.CompareTo(Christian) = IsGreater

Dim p1 As IObject
Dim p2 As IObject

Set p1 = cmd.Parameter(1)
Set p2 = p1

Debug.Print p1.Equals(p2)

Dim p3 As SqlParameter
Set p3 = SqlParameter.Build(TString.Build("Christian"), 1, "Name")
Debug.Print Christian.Equals(p3.Current)


Dim p4 As IObject
Set p4 = cmd.CreateParameter(TString.Build("04.04.20"), "Datum").AddValue(TString.Build("01.01.21"))

Debug.Print p4.Equals(p1)


End Sub
























Sub TestTType()

Dim n As TNumeric
Dim s As TString

Set n = TNumeric.Build(100.55, DefaultNumber)
Set s = TString.Build("test", , DefaultString)


End Sub

Sub TNumeric_Test()


Dim n As TNumeric
Set n = TNumeric.Build(2 ^ 31)

Debug.Print Skynet.IObject(n).HashValue

End Sub
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
    Dim v As IObject
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
