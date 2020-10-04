Attribute VB_Name = "Modul1"



'@Folder "Entwicklung"

Option Explicit

Private List As GenericList

Sub PrintSequence(ByVal E As IGenericIterator)

    With E
        Do While .MoveNext
            Debug.Print .Current
        Loop
    
        E.Reset
        
        Do While .MoveNext
            Debug.Print .Current
        Loop
    End With
    
End Sub

Sub TestArray()
    
    Dim T As CTimer
    Set T = New CTimer
    T.StartCounter
    
    Dim ga As GenericArray
    Dim ga2 As GenericArray
    Dim ga3 As GenericArray
    Dim x() As IGeneric
    Dim i As Long, n As Long
   
    n = 10000
    Set ga = GenericArray.Build(n)
    Set ga2 = GenericArray.Build(n)
    ReDim x(1 To n)
    
    For i = 1 To n
        Call ga.SetValue(GNumeric(i), i)
        Call ga2.SetValue(GString("Value: " & i), i)
        Set x(i) = GString("Value: " & i)
    Next
    
    Set ga3 = GenericArray.BuildFrom(x)
    
    Dim C As GenericArray
    Set C = Skynet.Clone(ga3)
    
    Dim Item As IGeneric

    With C.Iterator
        Do While .MoveNext
            Set Item = .Current
        Loop
    End With
    Debug.Print T.TimeElapsed
    
End Sub

Sub TestOrderedMap()
    
    Dim T As CTimer
    Set T = New CTimer
    T.StartCounter
    
    Dim map As GenericOrderedMap
    Set map = GenericOrderedMap.Build
    
    Dim i As Long, n As Long
   
    n = 10000
    For i = 1 To n
        Call map.Add(GNumeric(i), GNumeric(i))
    Next
    
    Dim C As GenericOrderedMap
    Set C = Skynet.Clone(map)
    
    Dim Item As IGeneric

    With C.Iterator(Pairs_)
        Do While .MoveNext
            Set Item = .Current
        Loop
    End With
    Debug.Print T.TimeElapsed
    
End Sub

Sub TestListIterator()

    Dim T As CTimer
    Set T = New CTimer
    T.StartCounter
    
    Dim List As GenericList
    Set List = GenericList.Build

    Dim i As Long, n As Long
    
    n = 10
    For i = 1 To n
        Call List.Add(GenericPair(GNumeric(i), GNumeric(i)))
    Next
    
    Dim C As GenericList
    Set C = Skynet.Clone(List)
    
    Dim Item As IGeneric
    With C.Iterator
        Do While .MoveNext
            Set Item = .Current
        Loop
    End With
    Debug.Print T.TimeElapsed

End Sub

Sub TestMaps()
    
    Dim T As CTimer
    
    Dim map As IGenericDictionary
    Set map = GenericTree.Build ' GenericOrderedMap.Build 'GenericSortedList.Build() 'GenericTree.Build '
    
    Dim i As Long, n As Long
    n = 1000
    
    If List Is Nothing Then
        Set List = GenericList.Build
        For i = 1 To n
            Call List.Add(GenericPair(GNumeric(i), GNumeric(i)))
        Next
        Call List.Sort(descending)
    End If
    
    Dim p As GenericPair
    Dim Item As IGeneric
    
    Set T = New CTimer
    T.StartCounter
   
    For i = 1 To n
        Set p = List(i)
        Call map.Add(p.Key, p.Value)
    Next
    Debug.Print T.TimeElapsed
    
    T.StartCounter
    For i = 1 To n
        Set p = List(i)
        Set Item = map.Item(p.Key)
    Next

    Debug.Print T.TimeElapsed
'
'    Dim ga As GenericArray
'    Set ga = GenericArray.Build(map.Count)
'    Call map.CopyOf(Pairs_, ga, ga.LowerBound)
'
'    For i = ga.LowerBound + 1 To ga.Length
'        If ga(i - 1).CompareTo(ga(i)) = IsGreater Then
'            Debug.Print "Error"
'        End If
'    Next

'    With map.Iterator(Pairs_)
'        Do While .MoveNext
'            Set Item = .Current
'        Loop
'    End With
'    Debug.Print T.TimeElapsed

End Sub


Sub TestSortedList()
    
    Dim T As CTimer
    Set T = New CTimer
    T.StartCounter
    
    Dim sl As GenericSortedList
    Set sl = GenericSortedList.Build()
    
    Dim i As Long, n As Long
    
    n = 1000
    For i = n To 1 Step -1
        Call sl.Add(GNumeric(i), GNumeric(i))
    Next
    Debug.Print T.TimeElapsed
    
    Dim C As GenericSortedList
    Set C = Skynet.Clone(sl)
    
    Dim Item As IGeneric

    With C.Iterator(Pairs_)
        Do While .MoveNext
            Set Item = .Current
        Loop
    End With
    Debug.Print T.TimeElapsed
    
End Sub

Sub TestTree()
    
    Dim T As CTimer
    Set T = New CTimer
    T.StartCounter
    
    Dim tree As GenericTree
    Set tree = GenericTree.Build
    
    Dim i As Long, n As Long

    n = 1000
    For i = n To 1 Step -1
        Call tree.Add(GNumeric(i), GNumeric(i))
    Next
    Debug.Print T.TimeElapsed
    
    Dim C As GenericTree
    Set C = Skynet.Clone(tree)
    
    Dim element As IGeneric
  
    With C.Iterator(Pairs_)
        Do While .MoveNext
            Set element = .Current
        Loop
    End With
    
    Debug.Print T.TimeElapsed
'
'    Dim collection As IGenericCollection
'    Set collection = C
'
'    Dim ga As GenericArray
'    Set ga = GenericArray.Build(collection.Count)
'    Call collection.CopyTo(ga, ga.LowerBound)
'    Call PrintSequence(ga.Iterator(5, 6))
'
End Sub

Sub TestArrayIterator()

    Dim i As Long
    Dim ga As GenericArray
    Set ga = GenericArray.Build(2)
    
    Set ga(1) = GString("A")
    Debug.Print ga(1)
    Call ga.SetValue(GString("b"), 2)
   Debug.Print ga.GetValue(2)
End Sub

Sub TestGenericCollection()
    
    Dim C As GenericOrderedMap
    Set C = GenericOrderedMap.Build
    
    Dim i As Long
    For i = 1 To 10
        Call C.Add(GString("Key: " & i), GString("Value: " & i))
    Next

    Dim List As GenericList
    Set List = GenericList.Build
    
    Call List.AddAll(C.GetValues)
   
    Dim Clone As IGenericReadOnlyList
    Set Clone = Skynet.Clone(List.AsReadOnly)
        
    Dim ga As GenericArray
    Set ga = GenericArray.Build(Clone.Count)
    
    Call Clone.CopyTo(ga, ga.LowerBound)
    Call Skynet.Dispose(Clone)
       
    For i = ga.LowerBound To ga.Length
        Debug.Print ga(i)
    Next

    Debug.Print ga.IndexOf(List(10))
    
    Call C.Remove(C.GetKeys(9))
    
End Sub
Sub TestGenericPair()
    
    Dim C As New VBA.collection
        
    Dim pair1 As IGeneric
    Set pair1 = GenericPair(GString("A"), Nothing)
    
    
    Dim pair2 As IGeneric
    Set pair2 = GenericPair(GString("A"), C)
    
    Debug.Print pair1.Equals(pair2)
End Sub

Sub asadsfsd()

Dim x(1 To 10) As IGeneric
Debug.Print Skynet.BinarySearch(x, Nothing, 1, 1, ascending)


End Sub

Sub assdsd()
    
    Dim i As Long
    Dim ga As GenericArray
    Set ga = GenericArray.Build(100)
    
    Call ga.SetValue(GString("b"), 13)
    Call ga.SetValue(GString("c"), 14)
    Call ga.SetValue(GString("a"), 15)
    Call ga.SetValue(GString("h"), 16)
    Call ga.SetValue(GString("s"), 17)
    Call ga.SetValue(GString("d"), 18)
    Call ga.SetValue(GString("zz"), 19)
    Call ga.SetValue(GString("c"), 20)
    Call ga.SetValue(GString("x"), 21)
    Call ga.SetValue(GString("e"), 22)
    Call ga.SetValue(GString("t"), 23)
    Call ga.SetValue(GString("a"), 24)

    Call ga.SetValue(GString("a"), 50)
    Call ga.SetValue(GString("c"), 51)
    Call ga.SetValue(GString("a"), 52)
    Call ga.SetValue(GString("j"), 53)
    Call ga.SetValue(GString("s"), 54)
    Call ga.SetValue(GString("ö"), 55)
    Call ga.SetValue(GString("q"), 56)
    Call ga.SetValue(GString("k"), 57)
    Call ga.SetValue(GString("x"), 58)
    Call ga.SetValue(GString("h"), 59)
    Call ga.SetValue(GString("t"), 60)
    Call ga.SetValue(GString("a"), 61)

    Call ga.SetValue(GString("z"), 70)
    Call ga.SetValue(GString("h"), 71)
    Call ga.SetValue(GString("t"), 72)
    Call ga.SetValue(GString("ä"), 73)

    Call ga.SetValue(GString("c"), 80)
    Call ga.SetValue(GString(""), 81)
    Call ga.SetValue(GString("e"), 82)
    Call ga.SetValue(GString("f"), 83)
    Call ga.SetValue(GString("d"), 84)
    Call ga.SetValue(GString("zz"), 85)
    Call ga.SetValue(GString("c"), 86)
    Call ga.SetValue(GString("x"), 87)
    Call ga.SetValue(GString("e"), 88)
    Call ga.SetValue(GString("f"), 89)
    Call ga.SetValue(GString("a"), 90)
    Call ga.SetValue(GString("a"), 100)
    Call ga.Sort(descending, ga.LowerBound, ga.Length)

    For i = 1 To ga.Length
        If Not ga(i) Is Nothing Then _
            Debug.Print "i: " & i & "  " & ga(i)
    Next
    
    Debug.Print ga.BinarySearch(GString("a"), 1, ga.Length, descending)
    Call ga.Reverse

End Sub

Public Sub Redim1()
    
    Dim i As Long
    ReDim a(1 To 50000) As IGeneric
    
    For i = 1 To 50000
        Set a(i) = GString("test" & i)
    Next
    Dim T As CTimer
    Set T = New CTimer
    T.StartCounter
    ReDim Preserve a(1 To 100000)
    Debug.Print T.TimeElapsed

End Sub

Public Sub Redim2()
    Dim T As CTimer
    Dim i As Long
'    Dim a(1 To 250) As IGeneric
'
    Dim a As GenericArray
    Set a = GenericArray.Build(250)

    For i = 1 To 250
        Set a(i) = GString("test" & i)
    Next
'
'    Set t = New Timer
'    t.StartCounter
'    For i = 1 To 250
'        Set a(i) = a(i)
'    Next
'    Debug.Print t.TimeElapsed

    Dim List As GenericList
    Set List = GenericList.Build()
    For i = 1 To 5
        List.Add GString("testX" & i)
    Next

    Call List.InsertAll(3, a)
    
    Call List.Reverse
End Sub

Public Sub normalArraytest()

    Dim i As Long
    ReDim a(10000) As IGeneric
    
    Dim T As CTimer
    Set T = New CTimer
    T.StartCounter
    
    For i = 1 To 10000
'        If i > UBound(a) Then _
'            ReDim Preserve a(UBound(a) * 2)
'
        Set a(i) = GString("Test1" & i)
    Next
    Debug.Print T.TimeElapsed

End Sub

Public Sub Arraytest()
    Dim i As Long
    Dim ga As GenericArray 'IGenericList
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
    
    Dim T As CTimer
    Set T = New CTimer
    
    Dim a(1 To 10000) As IGeneric
    Set ga = GenericArray.Build(10000)
    
    T.StartCounter
    For i = 1 To ga.Length
        Set ga(i) = GString("Test" & i)
        'GA(i).ToString
    Next
'    For i = 1 To GA.Length
'        Set GA(i) = GString("Test" & i)
'        'GA(i).ToString
'    Next
    Debug.Print T.TimeElapsed
     
'    t.StartCounter
'    For i = 1 To UBound(a)
'        Set a(i) = TString("Test" & i)
'        a(i).ToString
'    Next
'    Debug.Print t.TimeElapsed
    
    
    Dim ReadOnly As IGenericReadOnlyList

    
End Sub

Public Sub ListTest()

    Dim i As Long
    Dim List As GenericList
    Set List = GenericList.Build()
    
    Dim T As CTimer
    Set T = New CTimer
    T.StartCounter
    For i = 1 To 10000
        Call List.Add(GString("test" & i))
    Next
    Debug.Print T.TimeElapsed
'
'    Debug.Print List.IndexOf(GString("test" & 999), 1, 999)
'    Call List.Insert(500, GString("eingefügt an 500"))
'    Debug.Print List.IndexOf(GString("eingefügt an 500"))
'
    Dim List2 As GenericList
    Set List2 = Skynet.Clone(List)
    Debug.Print List2.Count
    Call List.Insert(500, GString("eingefügt an 500"))
    Debug.Print List(500)
    Debug.Print List2(500)
'
'    Dim List3 As GenericList
'    Set List3 = List.GetRange(1, 100)
'
''    For i = 1 To List2.Count
''        Debug.Print List2.ElementAt(i).ToString
''    Next
'    Dim readOnly As IGenericReadOnlyList
'    Set readOnly = List3.AsReadOnly
'    Debug.Print readOnly(1)
'    Set List = Nothing

End Sub

Public Sub Sometest()
    Dim i As Long
    
    Dim T As CTimer
    Set T = New CTimer
    
    Dim z(1 To 1000) As IGeneric

    For i = 1 To 1000
        Set z(i) = GString("test" & i)
    Next
    
    Dim x(1 To 1000) As IGeneric

    T.StartCounter
    For i = 1 To 1000
        Set x(i) = z(i)
    Next
    Debug.Print T.TimeElapsed
End Sub

Sub testMap()

    Dim i As Long
    Dim k As GString
    Dim v As GString
    
    Dim hm As GenericMap
    Set hm = GenericMap.Build()
    Dim T As CTimer
    Set T = New CTimer
    T.StartCounter
    For i = 1 To 2000
        Call hm.Add(GString("Key" & i), GNumeric(i), False)
    Next
    Debug.Print T.TimeElapsed

    Dim hm2 As GenericMap
    Set hm2 = Skynet.Clone(hm)
    Debug.Print Skynet.Generic(hm2).Equals(hm)
  
    'Call hm.Add(GString("Key" & 1), GString("ValueNew" & 1), False)'Error
    Debug.Print hm2.Item(GString("Key" & 1)).Equals(hm.Item(GString("Key" & 1)))
    
    Dim Values As GenericArray
    Set Values = hm.GetValues
    Call Values.Sort(-1)
    
'    With Keys.Iterator
'        Do While .MoveNext
'            Debug.Print .Current
'        Loop
'    End With
    
    i = Values.IndexOf(GNumeric(5))
    Debug.Print Values(i)
 
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


