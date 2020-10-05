Attribute VB_Name = "Modul1"

'@Folder "Entwicklung"

Option Explicit

Private List As GenericList

Sub TestArrayIterator()
   
    Dim i As Long, n As Long
    n = 50000
    
    Dim x() As IGeneric
    ReDim x(1 To n)
    
    For i = 1 To n
        Set x(i) = GNumeric(i)
    Next
    
    Dim Number As GNumeric
    Dim ga As GenericArray
    Set ga = GenericArray.BuildFrom(x)
     
    Dim T As CTimer
    Set T = New CTimer
    T.StartCounter
    For i = ga.LowerBound To ga.Length
        Set Number = ga(i)
    Next
    Debug.Print T.TimeElapsed
    
    T.StartCounter
    With ga.Iterator
        Do While .HasNext(Number) ' Fast
'            Set Number = .Current
        Loop
    End With
    Debug.Print T.TimeElapsed
    
End Sub
Sub TestRange()
    
    Dim Number As GNumeric
    Dim l As GenericList
    Set l = GenericList.Build
    
    Call l.AddAll(Skynet.Range(-5, 50))
    With l.Iterator
        Do While .HasNext(Number)
            Debug.Print Number.Value
        Loop
    End With
    
'    With Skynet.Range(-5, 5)
'        Do While .HasNext(Number)
'            Debug.Print Number.Value
'        Loop
'    End With

End Sub
Sub TestArraySort()
    Dim T As CTimer
    Set T = New CTimer
    
    Dim i As Long, n As Long
    n = 40000
    
    Set List = GenericList.Build(n)
    For i = 1 To n
'        Call List.Add(GenericPair(GNumeric(i), GNumeric(i)))
        Call List.Add(GNumeric(i))
    Next
    Call List.Sort(Random)
    
    T.StartCounter
    Call List.Sort(descending)
    Debug.Print T.TimeElapsed
    
    Dim Item As IGeneric
    With List.Iterator
        Do While .HasNext(Item)
          
        Loop
    End With
    
End Sub

Sub TestEquals()

    Dim T As CTimer
    Set T = New CTimer
    T.StartCounter
    
    Dim s1 As IGeneric
    Dim s2 As IGeneric
    
    Set s1 = GString("asasasfkfnkdfndjcfv falmxxs ejf")
    Set s2 = GString("asasasfkfnkdfndjcfv falmxxs ejf")
    
    Dim i As Long, n As Long
    
    T.StartCounter
    n = 10000
    For i = 1 To n
        s1.Equals s2
    Next
    Debug.Print T.TimeElapsed
     
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
        Do While .HasNext(Item)
            
        Loop
    End With
    Debug.Print T.TimeElapsed
    
End Sub

Sub TestOrderedMap()
    
    Dim T As CTimer
    Set T = New CTimer
    
    Dim Map As GenericOrderedMap
    Set Map = GenericOrderedMap.Build
    
    Dim Imap As IGenericDictionary
    Set Imap = GenericTree.Build
    
    Dim i As Long, n As Long
   
    n = 100
    For i = 1 To n
        Call Imap.Add(GNumeric(i), GNumeric(i))
    Next
    
    T.StartCounter
    Call Map.AddAll(Imap)
    Debug.Print T.TimeElapsed
    
    Dim C As GenericOrderedMap
    Set C = Skynet.Clone(Map)
    
    Dim Item As GenericPair

    With C.Iterator(Pairs_)
        Do While .HasNext(Item)
            
        Loop
    End With
    
End Sub

Sub TestListIterator()

    Dim T As CTimer
    Set T = New CTimer
    
    Dim List As GenericList
    Set List = GenericList.Build

    Dim i As Long, n As Long
    
    n = 10000
    For i = 1 To n
        Call List.Add(GenericPair(GNumeric(i), GNumeric(i)))
    Next
    
    Dim C As GenericList
    Set C = Skynet.Clone(List)
    
    T.StartCounter
    Dim Item As IGeneric
    With C.Iterator
        Do While .HasNext(Item)
           
        Loop
    End With
    Debug.Print T.TimeElapsed

End Sub

Sub TestMaps()
    
    Dim T As CTimer
    
    Dim Map As IGenericDictionary
    Set Map = GenericTree.Build ' GenericOrderedMap.Build 'GenericSortedList.Build() 'GenericTree.Build '
    
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
        Call Map.Add(p.Key, p.Value)
    Next
    Debug.Print T.TimeElapsed
    
    T.StartCounter
    For i = 1 To n
        Set p = List(i)
        Set Item = Map.Item(p.Key)
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
'        Do While .HasNext(Item)

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
        Do While .HasNext(Item)
           
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

    n = 10000
    For i = n To 1 Step -1
        Call tree.Add(GNumeric(i), GNumeric(i))
    Next
    Debug.Print T.TimeElapsed
    
    Dim C As GenericTree
    Set C = Skynet.Clone(tree)
    
    Dim Item As IGeneric
    
    With C.Iterator(Pairs_)
        Do While .HasNext(Item)
          
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
    
    Call List.AddAll(C.Iterator(Pairs_)) 'size is unknown
    'Call List.AddAll(C)' faster because size is known
   
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

End Sub

Sub TestArray2()
    
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

Sub testMap()

    Dim i As Long
    Dim K As GString
    Dim v As GString
    
    Dim hm As IGenericDictionary
    Set hm = GenericMap.Build()
    Dim T As CTimer
    Set T = New CTimer
    T.StartCounter
    For i = 1 To 20000
        Call hm.Add(GString("Key" & i), GNumeric(i))
    Next
    Debug.Print T.TimeElapsed
    
    Dim Item As IGeneric

    With GenericSortedList.Build(Dictionary:=hm).Iterator(Pairs_)
        Do While .HasNext(Item)
           
        Loop
    End With

End Sub

'Public Sub TestString()
'
'    Dim sentence As GString
'    Set sentence = GString("the quick brown fox jumps over the lazy dog")
'
'    Debug.Print "Before: " & sentence.Value
'
'    Set WordSequence = GenericEnumerable(sentence.Split(" ").Iterator(1, 9))
'
'    Debug.Print "After: " & WordSequence.Aggregate(GString, IgnoreNull:=True)
'
'End Sub
'
'Public Sub TestInteger()
'
'    Dim ints As GenericArray
'    Set ints = GenericArray.BuildWithIntegers(4, 8, 8, 3, 9, 0, 7, 8, 2)
'
'    Set IntegerSequence = GenericEnumerable(ints.Iterator(1, 9))
'    Debug.Print "Integers: " & GString.Join(ints, ";").Value
'
'    Debug.Print "The number of even integers is: " & IntegerSequence.Aggregate(GNumeric, IgnoreNull:=True)
'
'End Sub
'
'Private Sub IntegerSequence_Aggregate(Result As IGeneric, ByVal Current As IGeneric, ByVal NullsIgnored As Boolean)
'    If (Cast.ToNumeric(Current).IsEven) Then Set Result = Cast.ToNumeric(Result).Add(1)
'End Sub
'
'Private Sub WordSequence_Aggregate(Result As IGeneric, ByVal Current As IGeneric, ByVal NullsIgnored As Boolean)
'    Set Result = Cast.ToString(Current).Concat(Result, " ")
'End Sub


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


