Attribute VB_Name = "Modul1"

'@Folder "Entwicklung"

Option Explicit

Private SqlManager As GenericSqlManager
Private RandomList As GenericOrderedList

Private Type Table
    Name As Gstring
    
End Type

Sub TestInt()

    Dim Ints As GenericSortedSet
    
    Set Ints = GenericSortedSet.Of(IGenericComparer, Gint.Of(0), Gint.Of(-1), Gint.Of(5), Gint.Of(0), Gint.Of(1000))
    

End Sub
Sub TestBSearch()
    Dim List As GenericArray
    Set List = GenericArray.Of(Gnumeric(1), Gnumeric(2), Gnumeric(3), _
                                        Gnumeric(4), Gnumeric(5), Gnumeric(6), _
                                        Gnumeric(7), Gnumeric(8), Gnumeric(9))
                                        

    Debug.Print List.BinarySearch(Gnumeric(6), Ascending, IGenericComparer)

End Sub
Sub TestString()
    
    Dim Char As IGeneric
    
    Dim s As String
    Dim t As CTimer
    Set t = New CTimer
    Dim i As Long, n As Long
    n = 1000
    
    Dim Text As IGeneric, newText As Gstring
   
    t.StartCounter
    For i = 1 To n
        Set newText = Gstring.Of("abcdefghijklmnopqrstuvwxyz" & i)
'        Debug.Print newText.ElementAt(5).Value
    Next
    Debug.Print t.TimeElapsed
    
    Set Text = Gstring.Of("€tastatstastastsa" & i)
    
    t.StartCounter
    For i = 1 To n
       Call Text.Equals(newText)
    Next
    Debug.Print t.TimeElapsed
    
    t.StartCounter
    
    With Gstring.AsciiList.Sort.Elements.Iterator()
        Do While .HasNext(Char)
            Debug.Print Char.ToString
        Loop
    End With
  
    Debug.Print t.TimeElapsed
    
End Sub

Public Sub TestCollectionComparer()
    
    Dim Collection As IGenericCollection, Clone As IGenericCollection
    Set Collection = GenericOrderedList.Of(Gnumeric(0), Gnumeric(1), Gnumeric(2 ^ 36), Gstring("Test aoxdfidcisoxa,"))
    
    Dim Map As IGenericMap
    Set Map = GenericLinkedMap.Build(Comparer:=IGenericCollection)
    Call Map.TryAdd(Collection, Nothing)
    
    Set Clone = Collection.Copy
    
    Debug.Print Map.ContainsKey(Clone) = Map.ContainsKey(Collection)

End Sub
Public Sub TestSql()
    
    Dim t As CTimer
    Set t = New CTimer
    
'    Set Sql = GenericSqlManager.BuildSqlConnection(ServerName:="192.168.0.186", InitialCatalog:="TEST", User:="TestUser", Password:="OpenSesame")
    Set SqlManager = GenericSqlManager.BuildAccessConnection(Path:="C:\Daten\iCAT\Backend\Vers. 2.5\2020-02-24 iCAT-Backend Vers. 2.5.accdb", Filepassword:="OpenSesame")
    
    Dim Sql As Gstring
    Dim Table As Gstring
    Dim Columns As GenericOrderedList, Values As GenericOrderedList
    Dim Row As GenericOrderedMap
    
    Set Table = Gstring("tblG_00_Basis")
    
    t.StartCounter
    
    With SqlManager
        
        Set Columns = GenericOrderedList.Build.AddAll(.ColumnsOf(Table))
        Call Columns.Remove(Gstring("ID"))
        
        Set Sql = .SelectFrom(Table, Columns).Concat(.Where(Gstring("ID"), Gstring("=")), " ")
        
        With .Query(Sql, Columns, GenericArray.Of(Gnumeric(1)))
            Do While .HasNext(Row)
                With Row.Elements.Iterator 'Row.GetValues.Elements.Iterator
                    Do While .HasNext()
'                        Debug.Print .Current
                    Loop
                End With
            Loop
        End With
        
        Set Values = GenericOrderedList.Build.AddAll(Row.GetValues)
        Call Values.Add(Gnumeric(1))
        
        Set Sql = .Update(Table, Columns).Concat(.Where(Gstring("ID"), Gstring("=")), " ")
        Call .Execute(Sql, Row.GetValues)
            
        Debug.Print .IsConnected
        
    End With
    
'    Call Sql.Execute(CreateTables.Überblick)
'    Call Sql.Execute(CreateTables.Normal)
'    Call Sql.Execute(CreateTables.Intensiv)
'    Call Sql.Execute(CreateTables.Sanierung)
'    Call Sql.Execute(CreateTables.Abwicklung)
'    Call Sql.Execute(CreateTables.Paar)
'    Call Sql.Execute(CreateTables.Abstimmung)

End Sub

Sub CheckSql()

    Debug.Print SqlManager.IsConnected
End Sub
Sub TestMultiDimArray()

    Dim ga As GenericArray
    Set ga = GenericArray.Build(3, 4)
    
    Call ga.PutAt(Gnumeric(0), 0, 2)
    Call ga.PutAt(Gnumeric(1), 1, 2)
    Call ga.PutAt(Gnumeric(2), 2, 2)
    
    Dim Column As GenericArray
    Set Column = ga.SlizeColumn(2)
    
    Dim Element As Gnumeric
    With Column.Elements.Iterator
        Do While .HasNext(Element)
            Debug.Print Element.Value
        Loop
    End With
    
    'Insert/ Copy Column into Matrix first column
    Call Column.CopyTo(ga, ga.LowerBound)
    Debug.Print ga.GetAt(0, 0).Equals(ga.GetAt(0, 2))
    
    Call ga.Transpose
    Debug.Print ga.GetAt(2, 0).Equals(Column.ElementAt(0))
     Call Column.Elements.Clear
    
End Sub
Sub testArrayConstructor()
    
    Dim t As CTimer
    Set t = New CTimer
    
    Dim List As GenericOrderedList
    Dim i As Long
     
    t.StartCounter
   
    For i = 1 To 100
        Set List = GenericOrderedList.Of(Gstring, Gnumeric, Gnumeric)
    Next
    Debug.Print t.TimeElapsed
    
'    Dim Element As IGeneric
'    With List.Iterator
'        Do While .HasNext(Element)
'            Debug.Print Element
'        Loop
'    End With
'
End Sub
Sub TestArrayIterator2()
    
    Dim Char As IGeneric
    Dim s As Gstring
    Set s = Gstring("Ich bin ein Fuchs")
    
    With s.Chars.ToArray.Shuffle(12, 5).Elements.Iterator()
        Do While .HasNext(Char)
            Debug.Print Char
        Loop
    End With
    
    With s.Split(" ").Elements.Iterator
        Do While .HasNext(Char)
            Debug.Print Char
        Loop
    End With
    
    Debug.Print s.Chars.Contains(s.GetAt(0))
    Debug.Print s.Find("i")
    
End Sub
Sub TestArrayGetter()
    
    Dim t As CTimer
    Set t = New CTimer
    
    Dim List As GenericArray
    Dim Element As IGeneric
   
    Dim i As Long, n As Long
    n = 1000
        
    Set List = GenericArray.Build(n)
    ReDim X(1 To n) As IGeneric
    
    t.StartCounter
    For i = 1 To n
        Set List(i) = Gnumeric.Of(i)
    Next
    Debug.Print t.TimeElapsed
    
    t.StartCounter
    For i = 1 To n
        Set X(i) = Gint.Of(i)
    Next
    Debug.Print t.TimeElapsed
    
    
'
'    With List.Iterator
'        t.StartCounter
'        Do While .HasNext(Element)
'
'        Loop
'        Debug.Print t.TimeElapsed
'   End With

End Sub

Sub ListSort()
    Dim t As CTimer
    Set t = New CTimer
    
    Dim i As Long, n As Long
    n = 35
    
    Dim List As GenericOrderedList
    Set List = GenericOrderedList.Build(n)
   
    For i = 1 To n
'        Call List.Add(GenericPair(GNumeric(i), GNumeric(i)))
        Call List.Add(Gnumeric(i))
    Next
    Dim Item As IGeneric
    
    With List.Shuffle.Elements.Iterator()
        Do While .HasNext(Item)
            Debug.Print Item
        Loop
    End With
    
    t.StartCounter
    
    With List.Sort(Ascending).Elements.Iterator()
        Do While .HasNext(Item)
            Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed
End Sub

Sub TestEquals()

    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    
    Dim S1 As IGeneric
    Dim S2 As IGeneric
    
    Set S1 = Gstring("asasasfkfnkdfndjcfv falmxxs ejf")
    Set S2 = Gstring("asasasfkfnkdfndjcfv falmxxs ejf")
    
    Dim i As Long, n As Long
    
    t.StartCounter
    n = 10000
    For i = 1 To n
        S1.Equals S2
    Next
    Debug.Print t.TimeElapsed
     
End Sub

Sub TestListIterator()

    Dim t As CTimer
    Set t = New CTimer
    
    Dim l As GenericOrderedList
    Set l = GenericOrderedList.Build

    Dim i As Long, n As Long
    
    n = 15
    For i = 1 To n
        Call l.Add(GenericPair(Gnumeric(i), Gnumeric(i)))
    Next
    
    Dim c As GenericOrderedList
    Set c = System.Clone(l)
    
    t.StartCounter
    Dim Item As IGeneric
    With c.Elements.Iterator
        Do While .HasNext(Item)
           Debug.Print Item
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
        Call sl.Add(Gnumeric(i))
    Next
    
    Dim c As GenericSortedList
    Set c = System.Clone(sl)
    
    t.StartCounter
    Dim Item As IGeneric
    With c.Elements.Iterator
        Do While .HasNext(Item)
           Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed

End Sub

Sub TestSortedLists()
    Dim t As CTimer
    
    Dim P As GenericPair
    Dim Item As IGeneric
    
    Dim List As GenericSortedSet
    Set List = GenericSortedSet.Build(Comparer:=GenericPair)
    
    Dim i As Long, n As Long, j As Long
    n = 75
    
    If RandomList Is Nothing Then
        Set RandomList = GenericOrderedList.Build
        For i = 1 To n
            Call RandomList.Add(GenericPair(Gnumeric(i), Gnumeric(i)))
        Next
        Call RandomList.Shuffle
        
        With RandomList.Elements.Iterator
            Do While .HasNext(Item)
'                Debug.Print Item
            Loop
        End With
    End If

    Set t = New CTimer
    t.StartCounter
    For i = RandomList.First To RandomList.Last
        Set P = RandomList(i)
        Call List.Add(P)
    Next
    Debug.Print n & " elements added :: "; t.TimeElapsed

    For i = RandomList.First To RandomList.Last
        Set P = RandomList(i)
        If List.Contains(P) = False Then
            Debug.Print "not found in List :: " & P.Key
        End If
    Next
    
    For i = List.First + 1 To List.Last
        Debug.Print List.GetAt(i - 1)
        If Not List.Comparer.Compare(List.GetAt(i - 1), List.GetAt(i)) = islower Then
            Debug.Print "error"
        End If
    Next
    
'    t.StartCounter
'    With List.Elements.Iterator()
'        Do While .HasNext(Item)
'            Debug.Print Item
'        Loop
'    End With
'    Debug.Print t.TimeElapsed
'
End Sub
Sub TestSortedList2()
    
    Dim Value As IGeneric
    Dim Values As GenericOrderedList
    Set Values = GenericOrderedList.Of( _
                                            Gnumeric(12), _
                                            Gnumeric(56), _
                                            Gnumeric(1), _
                                            Gnumeric(67), _
                                            Gnumeric(45), _
                                            Gnumeric(8), _
                                            Gnumeric(82), _
                                            Gnumeric(16), _
                                            Gnumeric(63), _
                                            Gnumeric(23) _
                                         )
    'Call Values.Shuffle
    
    Dim i As Long, n As Long
    
    Dim sl As GenericSortedList
    Set sl = GenericSortedList.Build()
    Call sl.AddAll(Values)
    Call sl.AddAll(Values)
    
    With sl.Distinct.Elements.Iterator
'    With sl.Elements.Iterator
        Do While .HasNext(Value)
            Debug.Print Value
        Loop
    End With


End Sub
Sub TestSortedList()
    
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    
    Dim sl As GenericSortedList
    Dim i As Long, n As Long
    
    
    Set sl = GenericSortedList.Of(IGenericComparer, Gnumeric(7), Gnumeric(9), Gnumeric(3), Gnumeric(1), Gnumeric(10), Gnumeric(11), Gnumeric(13), Gnumeric(6), Gnumeric(8))
    
    Call sl.Add(Gnumeric(13))
  
    
    Call sl.AddAll(GenericArray.Of(Gnumeric(3), Gnumeric(4), Gnumeric(1), Gnumeric(5), Gnumeric(4), Gnumeric(2), Gnumeric(0), Gnumeric(12), Gnumeric(3)))
    
    Debug.Print t.TimeElapsed
    
    Dim c As GenericSortedList
    Set c = System.Clone(sl)
    
    Dim Item As IGeneric
    
    With c.Elements.Iterator
        Do While .HasNext(Item)
           Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed

End Sub

Sub TestTree()
    
    Dim i As Long
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    
    Dim tree As GenericSortedSet
    Set tree = GenericSortedSet.Build(Comparer:=IGenericComparer)
    Call tree.DoUnionWith(GenericArray.Of(Gnumeric(3), Gnumeric(4), Gnumeric(9), Gnumeric(1), Gnumeric(8), Gnumeric(5), Gnumeric(7), Gnumeric(0), Gnumeric(2), Gnumeric(0), Gnumeric(3), Gnumeric(10), Gnumeric(6), Gnumeric(8), Gnumeric(0)))
'    Call tree.DoExeceptWith(tree)

    For i = 0 To tree.Elements.Count - 1
        Debug.Print tree.GetAt(i)
    Next
    Debug.Print t.TimeElapsed
    
    Dim c As IGenericReadOnlyList
    Set c = System.Clone(tree)
    Call tree.Elements.Clear
    Set tree = Nothing
    
    Debug.Print c.IndexOf(Gnumeric(10))
    Debug.Print c.IndexOf(Gnumeric(1))
    Dim Item As IGeneric
    
    t.StartCounter
    
    With c.Elements.Iterator
        Do While .HasNext(Item)
            Debug.Print Item.ToString
        Loop
    End With
    
    Debug.Print t.TimeElapsed


End Sub

Sub TestGenericCollection()
    
    Dim c As GenericOrderedMap
    Set c = GenericOrderedMap.Build
    
    Dim Item As IGeneric
    
    Dim i As Long
    For i = 1 To 10
        Call c.Add(Gstring("Key: " & i), Gstring("Value: " & i))
    Next

    Dim List As GenericOrderedList
    Set List = GenericOrderedList.Build
    Call List.AddAll(c.Elements.Iterator()) 'size is unknown
    'Call List.AddAll(c)' faster because size is known
   
    Dim Clone As IGenericReadOnlyList
    Set Clone = List.Elements.Copy
        
    Dim ga As GenericArray
    Set ga = GenericArray.Build(Clone.Elements.Count)
    
    Call Clone.Elements.CopyTo(ga, ga.LowerBound) 'Set ga = Clone.Elements.ToArray
    
    Debug.Print ga.Elements.ContainsAll(List)
    
    For i = ga.LowerBound To ga.Length - 1
        Debug.Print ga.ElementAt(i)
    Next
    Set Item = List.ElementAt(List.Last)
    Debug.Print Item
    Debug.Print ga.IndexOf(Item)

End Sub

Sub TestArray2()
    
    Dim i As Long
    Dim ga As GenericArray
    Set ga = GenericArray.Build(100)
    Dim t As CTimer
    Set t = New CTimer
    With ga
        Call .PutAt(Gstring("b"), 13)
        Call .PutAt(Gstring("c"), 14)
        Call .PutAt(Gstring("a"), 15)
        Call .PutAt(Gstring("h"), 16)
        Call .PutAt(Gstring("s"), 17)
        Call .PutAt(Gstring("d"), 18)
        Call .PutAt(Gstring("zz"), 19)
        Call .PutAt(Gstring("c"), 20)
        Call .PutAt(Gstring("x"), 21)
        Call .PutAt(Gstring("e"), 22)
        Call .PutAt(Gstring("t"), 23)
        Call .PutAt(Gstring("a"), 24)
    
        Call .PutAt(Gstring("a"), 50)
        Call .PutAt(Gstring("c"), 51)
        Call .PutAt(Gstring("a"), 52)
        Call .PutAt(Gstring("j"), 53)
        Call .PutAt(Gstring("s"), 54)
        Call .PutAt(Gstring("ö"), 55)
        Call .PutAt(Gstring("q"), 56)
        Call .PutAt(Gstring("k"), 57)
        Call .PutAt(Gstring("x"), 58)
        Call .PutAt(Gstring("h"), 59)
        Call .PutAt(Gstring("t"), 60)
        Call .PutAt(Gstring("a"), 61)
    
        Call .PutAt(Gstring("z"), 70)
        Call .PutAt(Gstring("h"), 71)
        Call .PutAt(Gstring("t"), 72)
        Call .PutAt(Gstring("ä"), 73)
    
        Call .PutAt(Gstring("c"), 80)
        Call .PutAt(Gstring(""), 81)
        Call .PutAt(Gstring("e"), 82)
        Call .PutAt(Gstring("f"), 83)
        Call .PutAt(Gstring("d"), 84)
        Call .PutAt(Gstring("zz"), 85)
        Call .PutAt(Gstring("c"), 86)
        Call .PutAt(Gstring("x"), 87)
        Call .PutAt(Gstring("e"), 88)
        Call .PutAt(Gstring("f"), 89)
        Call .PutAt(Gstring("a"), 90)
        Call .PutAt(Gstring("a"), 99)
        
        t.StartCounter
        Call .Sort(Ascending, Comparer:=Nothing)
        Debug.Print t.TimeElapsed
'
'        For i = .LowerBound To .Length - 1
'            If Not ga.ElementAt(i) Is Nothing Then _
'                Debug.Print "i: " & i & "  " & ga.ElementAt(i)
'        Next

'        Debug.Print .BinarySearch(Gstring("zzz"), descending, .LowerBound, .Length, GenericValue.Comparer)
'        Call .Reverse
'        Call .Elements.Clear
    End With
    
End Sub

Public Sub TestList()

    Dim i As Long
    Dim Item As IGeneric
    Dim List As GenericOrderedList
    Set List = GenericOrderedList.Build()
    
    Dim t As CTimer
    Set t = New CTimer

    t.StartCounter
    Call List.Add(Gstring("test0"))
    Call List.AddAll(GenericArray.Of(Gstring("test1"), Gstring("test3"), Gstring("test2"), Gstring("test4"), Gstring("test5"), Gstring("test3")))
    With List.Elements.Iterator
        Do While .HasNext(Item)
            Debug.Print Item
        Loop
    End With
    
    Call List.RemoveAll(GenericArray.Of(Gstring("test3"), Gstring("test1")))
    
    With List.Elements.Iterator
        Do While .HasNext(Item)
            Debug.Print Item
        Loop
    End With
    
    Dim ArrayList As GenericArray
    Set ArrayList = GenericArray.Build(List.Elements.Count)

    Call List.CopyTo(ArrayList, ArrayList.LowerBound)
    Call List.Elements.Clear
    
    With ArrayList.Elements.Iterator
        Do While .HasNext(Item)
            Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed

End Sub

Sub testMap()

    Dim i As Long, Pair As IGeneric, Item As IGeneric
    
    Dim hm As GenericLinkedMap
    Set hm = GenericLinkedMap.Build(100)
    Dim t As CTimer
    Set t = New CTimer
    
    t.StartCounter
    For i = 1 To 40
        Call hm.TryAdd(Gstring("Key " & i), Gnumeric(i))
    Next
    Debug.Print t.TimeElapsed
    Debug.Print "Original"
    Debug.Print System.Generic(hm)
    
    For i = 1 To hm.Elements.Count

        Call hm.ContainsKey(Gstring("Key " & i))

    Next

    Dim Clone As IGenericMap
    Dim cmp As IGenericComparer
    Set cmp = IGenericCollection

    Set Clone = hm.Elements.Copy

    Debug.Print "Clone"
    Debug.Print System.Generic(Clone)

    Debug.Print "Deep equality :: " & cmp.Equals(hm, Clone)

    
    Debug.Print hm.ContainsKey(Nothing)
    Call hm.TryAdd(Nothing, Nothing)
    Debug.Print hm.ContainsKey(Nothing)
    
    With hm.Elements.Iterator
        Do While .HasNext(Pair)
            Debug.Print Pair
        Loop
    End With

    With hm.GetValues.Elements.Iterator()
        Do While .HasNext(Item)
'            Debug.Print Item.ToString
        Loop
    End With

    Set hm = Nothing
    Debug.Print t.TimeElapsed
    
End Sub

Sub testOrderedMap()

    Dim i As Long, Pair As IGeneric, Item As IGeneric
    
    Dim hm As GenericOrderedMap
    Set hm = GenericOrderedMap.Build(10)
    Dim t As CTimer
    Set t = New CTimer
    
    t.StartCounter
    For i = 1 To 30
        Call hm.Add(Gstring("Key " & i), Gnumeric(i))
    Next

    With hm.Elements.Iterator
        Do While .HasNext(Pair)
            Debug.Print Pair
        Loop
    End With

    With hm.GetValues.Elements.Iterator()
        Do While .HasNext(Item)
            Debug.Print Item.ToString
        Loop
    End With

    Set hm = Nothing
    Debug.Print t.TimeElapsed
    
End Sub


