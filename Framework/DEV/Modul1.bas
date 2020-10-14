Attribute VB_Name = "Modul1"

'@Folder "Entwicklung"

Option Explicit


Sub TestQuery()
    
    Dim t As New CTimer
    
    Dim Query As GenericQuery
    Dim Value As IGeneric
    Dim Name As GString
    Dim Name2 As GString
   
    Dim i As Long
    Dim n As Long
    
    n = 1000
    Dim Querys As GenericList
    Set Querys = GenericList.Build(n)
    
    t.StartCounter
    For i = 1 To n

         Set Query = New GenericQuery
         Call Query.Initialize(2)
         
         Set Name = GString("Parameter1")
         Set Name2 = GString("Parameter2")
         
         
'         Call Query.AddWithValues(Name, GenericList.BuildWith(GNumeric(100), GNumeric(1000), GNumeric(10000)))
'         Call Query.AddWithValues(Name2, GenericList.BuildWith(GString("value1"), GString("value2"), GString("value3")))

        Call Query.AddWithValues(Name, GNumeric(1000), GNumeric(100), GNumeric(10), GNumeric(1))
        Call Query.AddWithValues(Name2, GString("value3"), GString("value5"), GString("value2"), GString("value1"), GString("value4"))
     
        Call Querys.Add(Query)
    Next
    
    Dim SortedSet As GenericSortedSet
    Set Query = Querys(1)
    
    Set SortedSet = Query.GetValues(Name)
    Debug.Print SortedSet.ElementAt(1)
    Debug.Print t.TimeElapsed
   
    t.StartCounter
    Set Querys = Nothing
    Debug.Print t.TimeElapsed
    
    
End Sub
Sub TestCreation()
    
    Dim ga As GenericArray, Clone As GenericArray
    Dim i As Long, n As Long
   
    n = 10
    Set ga = GenericArray.Build(n)

    For i = 1 To n
        Call ga.SetValue(GNumeric(i), i)
    Next
    
    Set Clone = GenericArray.Build(ga.Length)
    Call GenericArray.Copy(ga, ga.LowerBound, Clone, Clone.LowerBound, ga.Length)
    
    Dim Element As IGeneric
    With Clone.Iterator
        Do While .HasNext(Element)
            Debug.Print Element
        Loop
    End With

End Sub

Sub TestMultiDimArray()

    Dim ga As GenericArray
    Set ga = GenericArray.Build(3, 4)
    
    Call ga.SetValue(GNumeric(1), 1, 3)
    Call ga.SetValue(GNumeric(2), 2, 3)
    Call ga.SetValue(GNumeric(3), 3, 3)
    
    Dim Column As GenericArray
    Set Column = ga.SlizeColumn(3)
    
    Dim Element As GNumeric
    With Column.Iterator
        Do While .HasNext(Element)
            Debug.Print Element.Value
        Loop
    End With
    
    'Insert/ Copy Column into Matrix first column
    Call GenericArray.Copy(Column, Column.LowerBound, ga, ga.LowerBound, Column.Length)
    Debug.Print ga.GetValue(1, 1).Equals(ga.GetValue(1, 3))
    
    Call ga.Transpose
    Debug.Print ga.GetValue(3, 1).Equals(Column(1))

End Sub
Sub testArrayConstructor()

    Dim ga As GenericArray
    Set ga = GenericArray.BuildWith(GNumeric(VBA.Now), GString("   now: " & VBA.Now & "!   ", Trim), GDate(VBA.Now, Year))
    
    Dim Element As IGeneric
    With ga.Iterator
        Do While .HasNext(Element)
            Debug.Print Element
        Loop
    End With
    
End Sub

Sub TestArrayIterator()
   
    Dim i As Long, n As Long
    n = 10000
    
    Dim x() As IGeneric
    ReDim x(1 To n)
    
    For i = 1 To n
        Set x(i) = GNumeric(i)
    Next
    
    Dim Number As GNumeric
    Dim ga As GenericArray
    Set ga = GenericArray.BuildFrom(x)
     
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    For i = ga.LowerBound To ga.Length
        Set Number = ga(i)
    Next
    Debug.Print t.TimeElapsed
    
    t.StartCounter
    With ga.Iterator
        Do While .HasNext(Number) ' Fast
'            Set Number = .Current
        Loop
    End With
    Debug.Print t.TimeElapsed
    
End Sub
Sub TestRange()
    
    Dim Number As GNumeric
    Dim L As GenericList
    Set L = GenericList.Build
    
    Call L.AddAll(Skynet.Range(-5, 25))
    With L.Iterator
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
    Dim t As CTimer
    Set t = New CTimer
    
    Dim i As Long, n As Long
    n = 40000
    
    Dim List As GenericList
    Set List = GenericList.Build(n)
    For i = 1 To n
'        Call List.Add(GenericPair(GNumeric(i), GNumeric(i)))
        Call List.Add(GNumeric(i))
    Next
    Call List.Sort(random)
    
    t.StartCounter
    Call List.Sort(descending)
    Debug.Print t.TimeElapsed
    
    Dim Item As IGeneric
    With List.Iterator
        Do While .HasNext(Item)
          
        Loop
    End With
    
End Sub

Sub TestEquals()

    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    
    Dim s1 As IGeneric
    Dim s2 As IGeneric
    
    Set s1 = GString("asasasfkfnkdfndjcfv falmxxs ejf")
    Set s2 = GString("asasasfkfnkdfndjcfv falmxxs ejf")
    
    Dim i As Long, n As Long
    
    t.StartCounter
    n = 10000
    For i = 1 To n
        s1.Equals s2
    Next
    Debug.Print t.TimeElapsed
     
End Sub
Sub TestArray()
    
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    
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
    
    Dim c As GenericArray
    Set c = Skynet.Clone(ga3)
    
    Dim Item As IGeneric
    
    With c.Iterator
        Do While .HasNext(Item)
            
        Loop
    End With
    Debug.Print t.TimeElapsed
    
End Sub

Sub TestOrderedMap()
    
    Dim t As CTimer
    Set t = New CTimer
    
    Dim Map As GenericOrderedMap
    Set Map = GenericOrderedMap.Build
    
    Dim Imap As IGenericDictionary
    Set Imap = GenericTree.Build
    
    Dim i As Long, n As Long
    
    n = 10000
    t.StartCounter
    For i = 1 To n
        Call Imap.Add(GNumeric(i), GNumeric(i))
    Next
    
    t.StartCounter
    Call Map.AddAll(Imap)
    Debug.Print t.TimeElapsed
    
    Dim c As GenericOrderedMap
    Set c = Skynet.Clone(Map)
    
    Dim Item As GenericPair
    t.StartCounter
    With c.Iterator(Pairs_)
        Do While .HasNext(Item)
'            Debug.Print Item.Key
        Loop
    End With
    Debug.Print t.TimeElapsed
    
End Sub

Sub TestListIterator()

    Dim t As CTimer
    Set t = New CTimer
    
    Dim L As GenericList
    Set L = GenericList.Build

    Dim i As Long, n As Long
    
    n = 5000
    For i = 1 To n
        Call L.Add(GenericPair(GNumeric(i), GNumeric(i)))
    Next
    
    Dim c As GenericList
    Set c = Skynet.Clone(L)
    
    t.StartCounter
    Dim Item As IGeneric
    With c.Iterator
        Do While .HasNext(Item)
'           Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed

End Sub

Sub TestSortedListIterator()

    Dim t As CTimer
    Set t = New CTimer
    
    Dim sl As GenericSortedList
    Set sl = GenericSortedList.Build

    Dim i As Long, n As Long
    
    n = 50
    For i = 1 To n
        Call sl.Add(GNumeric(i), GNumeric(i))
    Next
    
    Dim c As GenericSortedList
    Set c = Skynet.Clone(sl)
    
    t.StartCounter
    Dim Item As IGeneric
    With c.Iterator(t:=Keys_)
        Do While .HasNext(Item)
           Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed

End Sub

Sub TestMaps()
    
    Dim t As CTimer
    
    Dim Map As IGenericDictionary
    Set Map = GenericTree.Build ' GenericOrderedMap.Build 'GenericSortedList.Build() 'GenericTree.Build '
    
    Dim i As Long, n As Long, j As Long
    n = 1000
    
    Dim List As GenericList
    If List Is Nothing Then
        Set List = GenericList.Build
        For i = 1 To n
            Call List.Add(GenericPair(GNumeric(i), GNumeric(i)))
        Next
        Call List.Sort(random)
    End If
 
    Dim p As GenericPair
    Dim Item As IGeneric
    
    Set t = New CTimer
    t.StartCounter
    For i = 1 To n
        Set p = List(i)
        Call Map.Add(p.Key, p.Value)
    Next
    Debug.Print n & " :: "; t.TimeElapsed
  
    For i = 1 To n
        Set p = List(i)
        Set Item = Map.Item(p.Key)
    Next
    Dim Tree As GenericTree
    Set Tree = Map
    
    If Tree.Count = n = False Then
        Debug.Print "Tree.Count = n = False"
    Else
        Debug.Print Tree.ElementAt(n - 1)
    End If
    
    Set List = Nothing
    Set Item = Nothing
    Set Map = Nothing
    Set Tree = Nothing
'
'    Dim ga As GenericArray
'    Set ga = GenericArray.Build(Map.Count)
'    Call Map.CopyOf(Pairs_, ga, ga.LowerBound)
'
'    For i = ga.LowerBound + 1 To ga.Length
'        If ga(i - 1).CompareTo(ga(i)) = IsGreater Then
'            Debug.Print "Error"
'        End If
'    Next
''
'    t.StartCounter
'    With Map.Iterator(Pairs_)
'        Do While .HasNext(Item)
'            Debug.Print Item
'        Loop
'    End With
'    Debug.Print t.TimeElapsed
'
End Sub

Sub TestSortedList()
    
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    
    Dim sl As GenericSortedList
    Set sl = GenericSortedList.Build()
    
    Dim i As Long, n As Long
    
    n = 100
    For i = n To 1 Step -1
        Call sl.Add(GNumeric(i), GNumeric(i))
    Next
    Debug.Print t.TimeElapsed
    
    Dim c As GenericSortedList
    Set c = Skynet.Clone(sl)
    
    Dim Item As IGeneric

    With c.Iterator(Pairs_)
        Do While .HasNext(Item)
           
        Loop
    End With
    Debug.Print t.TimeElapsed

End Sub

Sub TestTree()
    
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    
    Dim Tree As GenericTree
    Set Tree = GenericTree.Build
    
    Dim i As Long
    
    For i = 1 To 100
        Call Tree.Add(GNumeric(i), Nothing)
    Next
  
    Dim c As GenericTree
    Set c = Skynet.Clone(Tree)
    
    Dim Item As IGeneric
    
    t.StartCounter
    With c.Iterator(Pairs_)
        Do While .HasNext(Item)
'            Debug.Print Item
        Loop
    End With
    
    Debug.Print t.TimeElapsed
'
'    Dim collection As IGenericCollection
'    Set collection = C
'
'    Dim ga As GenericArray
'    Set ga = GenericArray.Build(collection.Count)
'    Call collection.CopyTo(ga, ga.LowerBound)

End Sub

Sub TestGenericCollection()
    
    Dim c As GenericOrderedMap
    Set c = GenericOrderedMap.Build
    
    Dim i As Long
    For i = 1 To 10
        Call c.Add(GString("Key: " & i), GString("Value: " & i))
    Next

    Dim List As GenericList
    Set List = GenericList.Build
    
    Call List.AddAll(c.Iterator(Pairs_)) 'size is unknown
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
    
    With ga
        Call .SetValue(GString("b"), 13)
        Call .SetValue(GString("c"), 14)
        Call .SetValue(GString("a"), 15)
        Call .SetValue(GString("h"), 16)
        Call .SetValue(GString("s"), 17)
        Call .SetValue(GString("d"), 18)
        Call .SetValue(GString("zz"), 19)
        Call .SetValue(GString("c"), 20)
        Call .SetValue(GString("x"), 21)
        Call .SetValue(GString("e"), 22)
        Call .SetValue(GString("t"), 23)
        Call .SetValue(GString("a"), 24)
    
        Call .SetValue(GString("a"), 50)
        Call .SetValue(GString("c"), 51)
        Call .SetValue(GString("a"), 52)
        Call .SetValue(GString("j"), 53)
        Call .SetValue(GString("s"), 54)
        Call .SetValue(GString("ö"), 55)
        Call .SetValue(GString("q"), 56)
        Call .SetValue(GString("k"), 57)
        Call .SetValue(GString("x"), 58)
        Call .SetValue(GString("h"), 59)
        Call .SetValue(GString("t"), 60)
        Call .SetValue(GString("a"), 61)
    
        Call .SetValue(GString("z"), 70)
        Call .SetValue(GString("h"), 71)
        Call .SetValue(GString("t"), 72)
        Call .SetValue(GString("ä"), 73)
    
        Call .SetValue(GString("c"), 80)
        Call .SetValue(GString(""), 81)
        Call .SetValue(GString("e"), 82)
        Call .SetValue(GString("f"), 83)
        Call .SetValue(GString("d"), 84)
        Call .SetValue(GString("zz"), 85)
        Call .SetValue(GString("c"), 86)
        Call .SetValue(GString("x"), 87)
        Call .SetValue(GString("e"), 88)
        Call .SetValue(GString("f"), 89)
        Call .SetValue(GString("a"), 90)
        Call .SetValue(GString("a"), 100)
        Call .Sort(descending, ga.LowerBound, ga.Length)
    
        For i = 1 To .Length
            If Not ga(i) Is Nothing Then _
                Debug.Print "i: " & i & "  " & ga(i)
        Next
        
        Debug.Print .BinarySearch(GString("a"), 1, .Length, descending)
        Call .Reverse
        Call .Clear
    End With
    
End Sub

Public Sub ListTest()

    Dim i As Long
    Dim List As GenericList
    Set List = GenericList.Build()
    
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    For i = 1 To 1000
        Call List.Add(GString("test" & i))
    Next
    Debug.Print t.TimeElapsed
'
    Debug.Print List.IndexOf(GString("test" & 999), 1, 999)
    Call List.Insert(500, GString("eingefügt an 500"))
    Debug.Print List.IndexOf(GString("eingefügt an 500"))
'
    Dim List2 As GenericList
    Set List2 = Skynet.Clone(List)
    Debug.Print List2.Count
    Call List.Insert(500, GString("eingefügt an 500"))
    Debug.Print List(500)
    Debug.Print List2(500)

    Dim List3 As GenericList
    Set List3 = List.GetRange(500, 503)
    
    Dim readOnly As IGenericReadOnlyList
    Set readOnly = List3.AsReadOnly
    Debug.Print readOnly(1)
    Debug.Print readOnly(10)
    Debug.Print readOnly.Count
    Set List = Nothing

End Sub

Sub testMap()

    Dim i As Long
 
    Dim hm As IGenericDictionary
    Set hm = GenericMap.Build()
    Dim t As CTimer
    Set t = New CTimer
   
    For i = 1 To 10
        Call hm.Add(GString("Key" & i), GNumeric(i))
    Next
    
    Dim Clone As IGenericDictionary
    Set Clone = Skynet.Clone(hm)
    Set hm = Nothing
    
    Dim Item As IGeneric
    t.StartCounter
    With Clone.Iterator(t:=Pairs_)
        Do While .HasNext(Item)
'            Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed
    
    With GenericSortedList.BuildFrom(Dictionary:=Clone).Iterator(Pairs_)
        Do While .HasNext(Item)
            Debug.Print Item
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


