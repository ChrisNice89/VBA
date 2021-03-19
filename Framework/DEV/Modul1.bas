Attribute VB_Name = "Modul1"

'@Folder "Entwicklung"

Option Explicit

Private Sql As GenericSqlManager
Private RandomList As GenericOrderedList

Private Type Table
    Name As Gstring
    
End Type

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
        Set newText = Gstring.Build("abcdefghijklmnopqrstuvwxyz" & i)
'        Debug.Print newText.ElementAt(5).Value
    Next
    Debug.Print t.TimeElapsed
    
    Set Text = Gstring.Build("€tastatstastastsa" & i)
    
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
    Set Collection = GenericOrderedList.AsList(Gnumeric(0), Gnumeric(1), Gnumeric(2 ^ 36), Gstring("Test aoxdfidcisoxa,"))
    
    Dim Map As IGenericMap
    Set Map = GenericLinkedMap.Build(Comparer:=IGenericCollection)
    Call Map.Add(Collection, Nothing)
    
    Set Clone = Collection.Copy
    
    Debug.Print Map.ContainsKey(Clone) = Map.ContainsKey(Collection)

End Sub
Public Sub TestSql()
    
    Dim t As CTimer
    Set t = New CTimer
    
'    Set Sql = GenericSqlManager.BuildSqlConnection(ServerName:="192.168.0.186", InitialCatalog:="TEST", User:="TestUser", Password:="OpenSesame")
    Set Sql = GenericSqlManager.BuildAccessConnection(Path:="C:\Daten\iCAT\Backend\Vers. 2.5\2020-02-24 iCAT-Backend Vers. 2.5.accdb", Filepassword:="OpenSesame")
    
    Dim Table As Gstring
    Dim Columns As IGenericReadOnlyList
    Dim Where As GenericPair, Operator As Gstring
    
    Dim Field As IGeneric
    Dim Updates As IGenericMap
    Dim Row As GenericOrderedMap
    
    Set Table = Gstring("tblG_00_Basis")
    Set Where = GenericPair(Gstring("ID"), Gnumeric(1))
    Set Operator = Gstring("=")
    Set Columns = Sql.ColumnsOf(Table)
    
    t.StartCounter
    With Sql.SelectWhere(Table, Columns, Where, Operator)
    Debug.Print t.TimeElapsed
        Do While .HasNext(Row)
            With Row.Elements.Iterator 'Row.GetValues.Elements.Iterator
                Do While .HasNext(Field)
'                    Debug.Print Field
                Loop
            End With
        Loop
    End With
    
    Debug.Print Sql.IsConnected
    
    
'    Call Sql.UpdateWhere(Table, Row, Where, Operator)
    
    
'    Call Sql.Execute(CreateTables.Überblick)
'    Call Sql.Execute(CreateTables.Normal)
'    Call Sql.Execute(CreateTables.Intensiv)
'    Call Sql.Execute(CreateTables.Sanierung)
'    Call Sql.Execute(CreateTables.Abwicklung)
'    Call Sql.Execute(CreateTables.Paar)
'    Call Sql.Execute(CreateTables.Abstimmung)

End Sub

Sub CheckSql()

    Debug.Print Sql.IsConnected
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
        Set List = GenericOrderedList.AsList(Gstring, Gnumeric, Gnumeric)
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
    ReDim x(1 To n) As IGeneric
    
    t.StartCounter
    For i = 1 To n
        Set List(i) = Gnumeric.Build(i)
    Next
    Debug.Print t.TimeElapsed
    
    t.StartCounter
    For i = 1 To n
        Set x(i) = Gnumeric.Build(i)
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
    
    With List.Sort(ascending).Elements.Iterator()
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
    
    Dim C As GenericOrderedList
    Set C = System.Clone(l)
    
    t.StartCounter
    Dim Item As IGeneric
    With C.Elements.Iterator
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
    
    Dim C As GenericSortedList
    Set C = System.Clone(sl)
    
    t.StartCounter
    Dim Item As IGeneric
    With C.Elements.Iterator
        Do While .HasNext(Item)
           Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed

End Sub

Sub TestSortedLists()
    Dim t As CTimer
    
    Dim p As GenericPair
    Dim Item As IGeneric
    
    Dim List As GenericSortedList
    Set List = GenericSortedList.Build(Comparer:=GenericPair)
    
    Dim i As Long, n As Long, j As Long
    n = 30
    
    If RandomList Is Nothing Then
        Set RandomList = GenericOrderedList.Build
        For i = 1 To n
            Call RandomList.Add(GenericPair(Gnumeric(i), Gnumeric(i)))
        Next
        Call RandomList.Shuffle
        
        With RandomList.Elements.Iterator
            Do While .HasNext(Item)
                Debug.Print Item
            Loop
        End With
    End If

    Set t = New CTimer
    t.StartCounter
    For i = RandomList.First To RandomList.Last
        Set p = RandomList(i)
        Call List.Add(p)
    Next
    Debug.Print n & " elements added :: "; t.TimeElapsed

    For i = RandomList.First To RandomList.Last
        Set p = RandomList(i)
        If List.Contains(p) = False Then
            Debug.Print "not found in List :: " & p.Key
        End If
    Next
  
    t.StartCounter
    With List.Elements.Iterator()
        Do While .HasNext(Item)
            Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed
'
End Sub
Sub TestSortedList2()
    
    Dim Value As IGeneric
    Dim Values As GenericOrderedList
    Set Values = GenericOrderedList.AsList( _
                                            Gnumeric(1), _
                                            Gnumeric(3), _
                                            Gnumeric(5), _
                                            Gnumeric(7), _
                                            Gnumeric(9), _
                                            Gnumeric(11), _
                                            Gnumeric(13), _
                                            Gnumeric(15), _
                                            Gnumeric(17), _
                                            Gnumeric(19), _
                                            Gnumeric(21))
    Call Values.Shuffle
    
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
    Debug.Print sl.Contains(Gnumeric(21))
    

End Sub
Sub TestSortedList()
    
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    
    Dim sl As GenericSortedList
    Dim i As Long, n As Long
    
    
    Set sl = GenericSortedList.AsList(IGenericComparer, Gnumeric(7), Gnumeric(9), Gnumeric(3), Gnumeric(1), Gnumeric(10), Gnumeric(11), Gnumeric(13), Gnumeric(6), Gnumeric(8))
    
    Call sl.Add(Gnumeric(13))
  
    
    Call sl.AddAll(GenericArray.AsArray(Gnumeric(3), Gnumeric(4), Gnumeric(1), Gnumeric(5), Gnumeric(4), Gnumeric(2), Gnumeric(0), Gnumeric(12), Gnumeric(3)))
    
    Debug.Print t.TimeElapsed
    
    Dim C As GenericSortedList
    Set C = System.Clone(sl)
    
    Dim Item As IGeneric
    
    With C.Elements.Iterator
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
    Call tree.DoUnionWith(GenericArray.AsArray(Gnumeric(3), Gnumeric(4), Gnumeric(1), Gnumeric(5), Gnumeric(7), Gnumeric(0), Gnumeric(2), Gnumeric(0), Gnumeric(3), Gnumeric(10)))
'    Call tree.DoExeceptWith(tree)

    For i = 0 To tree.Elements.Count - 1
        Debug.Print tree.GetAt(i)
    Next
    Debug.Print t.TimeElapsed
    
    Dim A As GenericArray
'    Set a = GenericArray.Build(30)
'
'    Call tree.CopyTo(a, 0, 6, 8)
'
'    Dim N As IGeneric
'    Set N = tree.ElementAt(1)
'    Debug.Print N.ToString
    Dim C As IGenericReadOnlyList
    Set C = System.Clone(tree)
    Call tree.Elements.Clear
    Set tree = Nothing
    
    Debug.Print C.IndexOf(Gnumeric(10))
    Debug.Print C.IndexOf(Gnumeric(1))
    Dim Item As IGeneric
    
    t.StartCounter
    
    With C.Elements.Iterator
        Do While .HasNext(Item)
            Debug.Print Item.ToString
        Loop
    End With
    
    Debug.Print t.TimeElapsed


End Sub

Sub TestGenericCollection()
    
    Dim C As GenericOrderedMap
    Set C = GenericOrderedMap.Build
    
    Dim Item As IGeneric
    
    Dim i As Long
    For i = 1 To 10
        Call C.Add(Gstring("Key: " & i), Gstring("Value: " & i))
    Next

    Dim List As GenericOrderedList
    Set List = GenericOrderedList.Build
    Call List.AddAll(C.Elements.Iterator()) 'size is unknown
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
        Call .Sort(descending, Comparer:=Nothing)
        Debug.Print t.TimeElapsed
        
        For i = .LowerBound To .Length - 1
            If Not ga.ElementAt(i) Is Nothing Then _
                Debug.Print "i: " & i & "  " & ga.ElementAt(i)
        Next

        Debug.Print .BinarySearch(Gstring("zzz"), descending, .LowerBound, .Length, GenericValue.Comparer)
        Call .Reverse
        Call .Elements.Clear
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
    Call List.AddAll(GenericArray.AsArray(Gstring("test1"), Gstring("test3"), Gstring("test2"), Gstring("test4"), Gstring("test5"), Gstring("test3")))
    With List.Elements.Iterator
        Do While .HasNext(Item)
            Debug.Print Item
        Loop
    End With
    
    Call List.RemoveAll(GenericArray.AsArray(Gstring("test3"), Gstring("test1")))
    
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
    For i = 1 To 5000
        Call hm.Add(Gstring("Key " & i), Gnumeric(i))
    Next
    Debug.Print t.TimeElapsed
'    Debug.Print "Original"
'    Debug.Print System.Generic(hm)
    
'    For i = 1 To hm.Elements.Count
'
'        Call hm.ContainsKey(GString("Key " & i))
'
'    Next
'
'    Dim Clone As IGenericMap
'    Dim cmp As IGenericComparer
'    Set cmp = IGenericCollection
'
'    Set Clone = hm.Elements.Copy
'
'    Debug.Print "Clone"
'    Debug.Print System.Generic(Clone)
'
'    Debug.Print "Deep equality :: " & cmp.Equals(hm, Clone)
    Exit Sub
    
'    Debug.Print hm.ContainsKey(Nothing)
'    Call hm.Add(Nothing, Nothing)
'    Debug.Print hm.ContainsKey(Nothing)
    
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
'            Debug.Print Item.ToString
        Loop
    End With

    Set hm = Nothing
    Debug.Print t.TimeElapsed
    
End Sub


