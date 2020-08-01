Attribute VB_Name = "Modul1"
'@Folder "Entwicklung"

Option Explicit
Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long

Public Sub CollectionTest()

    Dim x As IGeneric
    
    Dim i As Long
    Dim List As Collection
    Set List = New Collection
    
    Dim t As Timer
    Set t = New Timer
  
    t.StartCounter
    For i = 1 To 50000
        Call List.Add(GString("test" & i))

    Next
    Debug.Print t.TimeElapsed

End Sub

Public Sub normalArraytest()

    Dim i As Long
    ReDim a(10000) As IGeneric
    
    Dim t As Timer
    Set t = New Timer
    t.StartCounter
    
    For i = 1 To 10000
'        If i > UBound(a) Then _
'            ReDim Preserve a(UBound(a) * 2)
'
        Set a(i) = GString("Test1" & i)
    Next
    Debug.Print t.TimeElapsed

End Sub

Public Sub Arraytest()
    Dim i As Long
    Dim GA As GenericArray 'IGenericList
'    Set sa = SafeArray.Build(10, 2)
'
'    For i = 1 To sa.Count
'        Set sa.Object(i, 1) = TString("Test1" & i)
'        Set sa.Object(i, 2) = TString("Test2" & i)
'    Next
'
'    Dim field As SafeArray
'    Set field = sa.SlizeColumn(2)
'
'    Dim b As SafeArray
'    Set b = SafeArray.Build(9)
'
'    Call sa.CopyTo(b, b.AdressOf(9), sa.AdressOf(9, 2), 1)
'
'    Call sa.Transpose
    
    Dim t As Timer
    Set t = New Timer
    
    Dim a(1 To 10000) As IGeneric
    Set GA = GenericArray.Build(10000)
    
    t.StartCounter
    For i = 1 To GA.Length
        Set GA(i) = GString("Test" & i)
        'GA(i).ToString
    Next
'    For i = 1 To GA.Length
'        Set GA(i) = GString("Test" & i)
'        'GA(i).ToString
'    Next
    Debug.Print t.TimeElapsed
     
'    t.StartCounter
'    For i = 1 To UBound(a)
'        Set a(i) = TString("Test" & i)
'        a(i).ToString
'    Next
'    Debug.Print t.TimeElapsed
    
End Sub

Public Sub ListTest()

    Dim i As Long
    Dim List As GenericList
    Set List = GenericList.Build(1000)
    
    For i = 1 To 1000
        Call List.Add(GString("test" & i))
    Next
    Debug.Print List.IndexOf(GString("test" & 999), 999, 2)
    Call List.Insert(500, GString("eingefügt an 500"))
    Debug.Print List(500)
    
    Dim List2 As GenericList
    Set List2 = Skynet.Generic(List).Clone
    Debug.Print List2.Count
    Call List.Insert(1, GString("eingefügt an 1"))
    Debug.Print List(1)
    Debug.Print List2(1)
    
    Dim List3 As GenericList
    Set List3 = List.SubList(1002, 1002)

'    For i = 1 To List2.Count
'        Debug.Print List2.ElementAt(i).ToString
'    Next
    Dim roList As GenericReadOnlyList
    Set roList = List3.AsReadOnly
    Debug.Print roList(1)
    Set List = Nothing

End Sub

Public Sub Sometest()
    Dim i As Long
    
    Dim t As Timer
    Set t = New Timer
    
    Dim z(1 To 1000) As IGeneric

    For i = 1 To 1000
        Set z(i) = GString("test" & i)
    Next
    
    Dim x(1 To 1000) As IGeneric

    t.StartCounter
    For i = 1 To 1000
        Set x(i) = z(i)
    Next
    Debug.Print t.TimeElapsed
End Sub


Sub TMaptest()

    Dim i As Long
    Dim K As GString
    Dim V As GString
    
    Dim hm As GenericMap
    Set hm = GenericMap.Build(50000)
    
    For i = 1 To 1000
        Call hm.Add(GString("Key" & i), GString("Value" & i)) 'TString("Value" & i)
    Next
    Debug.Print hm(GString("Key1"))
    Dim hm2 As GenericMap
    Set hm2 = Skynet.Generic(hm).Clone
    
    Debug.Print hm2 Is hm
    Debug.Print Skynet.Generic(hm2).Equals(hm)
    Debug.Print hm2.Item(GString("Key" & 1)).ToString
    
    Call hm.Add(GString("Key" & 1), GString("ValueNew" & 1))
    Debug.Print hm2.Item(GString("Key" & 1)).ToString
    Debug.Print hm.Item(GString("Key" & 1)).ToString
    
End Sub
'Sub Cmdtest()
'
'    Dim cmd As SqlCommand
'    Dim Sql As String
'    Sql = "SomeSql"
'    Set cmd = SqlCommand.Build(Sql, SqlConnection)
'    Debug.Print cmd.Sql.Replace("Some", "Somee").Value
'End Sub
'
'Sub ParameterTest()
'
'Dim cmd As SqlCommand
'Set cmd = SqlCommand.Build("SomeSql", SqlConnection)
'
'Call cmd.CreateParameter(GString("Christian"), "Name").AddValue(GString("Christoph"))
'Debug.Print cmd.Parameter("Name").CurrentValue.Value
'
'Debug.Print cmd.Parameter("Name").UseValue(2).Object.ToString
'
'Dim Christian As IGeneric
'Dim Christoph As IGeneric
'
'Set Christian = cmd.Parameter("Name").Value(1)
'Set Christoph = cmd.Parameter("Name").Value(2)
'
'Debug.Print Christoph.CompareTo(Christian) = IsGreater
'
'Dim p1 As IGeneric
'Dim p2 As IGeneric
'
'Set p1 = cmd.Parameter(1)
'Set p2 = p1
'
'Debug.Print p1.Equals(p2)
'
'Dim p3 As SqlParameter
'Set p3 = SqlParameter(GString("Christian"), "Name")
'Debug.Print Christian.Equals(p3.CurrentValue)
'
'
'Dim p4 As IGeneric
''Set p4 = cmd.CreateParameter(TDate(#4/4/2020#), "Datum").AddValue(TDate(#1/1/2021#))
'
'Debug.Print p4.Equals(p1)
'
'
'End Sub
'
'


