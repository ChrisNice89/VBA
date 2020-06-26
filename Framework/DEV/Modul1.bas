Attribute VB_Name = "Modul1"
'@Folder "Entwicklung"

Option Explicit


Sub TmapTest()

    Dim i As Long
    Dim K As TString
    Dim v As TString
    
    Dim hm As TMap
    Dim T As Timer
    Set T = New Timer
  
    Set hm = TMap.Build()
    
    T.StartCounter
    For i = 1 To 10000
        Set K = TString("Key" & i)
        Set v = TString("Value" & i)
        Call hm.Add(K, v)
    Next
    Debug.Print T.TimeElapsed
    
End Sub
Sub Cmdtest()

    Dim cmd As SqlCommand
    Dim Sql As String
    Sql = "SomeSql"
    Set cmd = SqlCommand.Build(Sql, SqlConnection)
    Debug.Print cmd.Sql.Replace("Some", "Somee").Value
End Sub

Sub ParameterTest()

Dim cmd As SqlCommand
Set cmd = SqlCommand.Build("SomeSql", SqlConnection)

Call cmd.CreateParameter(TString("Christian"), "Name").AddValue(TString("Christoph"))
Debug.Print cmd.Parameter("Name").CurrentValue.Value

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
Set p3 = SqlParameter(TString("Christian"), "Name")
Debug.Print Christian.Equals(p3.CurrentValue)


Dim p4 As IObject
'Set p4 = cmd.CreateParameter(TDate(#4/4/2020#), "Datum").AddValue(TDate(#1/1/2021#))

Debug.Print p4.Equals(p1)


End Sub
















Sub TestTType()

Dim n As TNumeric
Dim s As TString

Set n = TNumeric.Build(100.55, DefaultNumber)
Set s = TString("test", , DefaultString)


End Sub

Sub TNumeric_Test()


Dim n As TNumeric
Set n = TNumeric.Build(2 ^ 31)

Debug.Print Skynet.CastObject(n).HashValue

End Sub
Sub Compare()


    Dim K As TString
    Dim v As TString
    Dim i As Long
    Dim A(1 To 1000) As IObject
    Dim B(1 To 1000) As IObject
    
    For i = 1 To 1000
        Set A(i) = TString("A")
        Set B(i) = TString("B")
    Next
    
    Dim T As New Timer
    T.StartCounter
    For i = 1 To 1000
        If Not A(i).CompareTo(B(i)) = IsLower Then
            Debug.Print "error"
        End If
    Next
    Debug.Print T.TimeElapsed



End Sub

Sub Equals()

    Dim K As TString
    Dim v As TString
    Dim i As Long
    Dim A(1 To 1000) As IObject
    Dim B(1 To 1000) As IObject
    
    For i = 1 To 1000
        Set A(i) = TString("Key" & i)
        Set B(i) = TString("xey" & i)
    Next
    
    Dim T As New Timer
    T.StartCounter
    For i = 1 To 1000
        Call A(i).Equals(B(i))
    Next
    Debug.Print T.TimeElapsed

End Sub

Sub test()

    Dim K As TString
    Dim v As TString
    Dim i As Long
    
    Dim map As TMap
    Set map = TMap.Build '(TString, TString)
    Dim T As New Timer
    T.StartCounter
    For i = 1 To 10000
        Set K = TString("Key" & i)
        Set v = TString("Value" & i)
        Call map.Add(K, v)
    Next
    Debug.Print T.TimeElapsed
    Debug.Print Skynet.CastObject(map).ToString
End Sub

Sub ttt()
    Dim s As TString
    Dim n As TNumeric
    Dim D As TDate
    Dim B As TBoolean
    Dim f As TFloat
    Dim i As Long
    Dim T As New Timer
    
    T.StartCounter
    For i = 1 To 10000
        Set s = TString("Value" & i)
        Set n = TNumeric(i)
        Set D = TDate(#1/1/2020#)
        Set B = TBoolean(True)
        Set f = TFloat(i / 100)
        
    Next
    Debug.Print T.TimeElapsed

End Sub

Public Sub Is64BitExcel()


#If VBA7 And Win64 Then
    Debug.Print True
#Else
    Debug.Print False
#End If
End Sub
