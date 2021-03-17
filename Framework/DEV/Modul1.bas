Attribute VB_Name = "Modul1"

'@Folder "Entwicklung"

Option Explicit

Private Sql As GenericSqlManager
Private RandomList As GenericOrderedList

Private Type Table
    Name As GString
    
End Type

Sub TestString()
    
    Dim Char As IGeneric
    
    Dim s As String
    Dim t As CTimer
    Set t = New CTimer
    Dim i As Long, N As Long
    N = 1000
    
    Dim Text As IGeneric, newText As GString
   
    t.StartCounter
    For i = 1 To N
        Set newText = GString.Build("abcdefghijklmnopqrstuvwxyz" & i)
'        Debug.Print newText.ElementAt(5).Value
    Next
    Debug.Print t.TimeElapsed
    
    Set Text = GString.Build("€tastatstastastsa" & i)
    
    t.StartCounter
    For i = 1 To N
       Call Text.Equals(newText)
    Next
    Debug.Print t.TimeElapsed
    
    t.StartCounter
    
    With GString.AsciiList.Sort.Elements.Iterator()
        Do While .HasNext(Char)
            Debug.Print Char.ToString
        Loop
    End With
  
    Debug.Print t.TimeElapsed
    
End Sub

Public Sub TestCollectionComparer()
    
    Dim Collection As IGenericCollection, Clone As IGenericCollection
    Set Collection = GenericOrderedList.AsList(GNumeric(0), GNumeric(1), GNumeric(2 ^ 36), GString("Test aoxdfidcisoxa,"))
    
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
    
    Dim Table As GString
    Dim Columns As IGenericReadOnlyList
    Dim Where As GenericPair, Operator As GString
    
    Dim Field As IGeneric
    Dim Updates As IGenericMap
    Dim Row As GenericOrderedMap
    
    Set Table = GString("tblG_00_Basis")
    Set Where = GenericPair(GString("ID"), GNumeric(1))
    Set Operator = GString("=")
    Set Columns = Sql.ColumnsOf(Table)
    
    t.StartCounter
    With Sql.SelectWhere(Table, Columns, Where, Operator)
        Do While .HasNext(Row)
            With Row.Elements.Iterator 'Row.GetValues.Elements.Iterator
                Do While .HasNext(Field)
                    Debug.Print Field
                Loop
            End With
        Loop
    End With
    Debug.Print t.TimeElapsed
    
    
    Call Sql.UpdateWhere(Table, Row, Where, Operator)
    
    
'    Call Sql.Execute(CreateTables.Überblick)
'    Call Sql.Execute(CreateTables.Normal)
'    Call Sql.Execute(CreateTables.Intensiv)
'    Call Sql.Execute(CreateTables.Sanierung)
'    Call Sql.Execute(CreateTables.Abwicklung)
'    Call Sql.Execute(CreateTables.Paar)
'    Call Sql.Execute(CreateTables.Abstimmung)
    
End Sub

Sub TestMultiDimArray()

    Dim ga As GenericArray
    Set ga = GenericArray.Build(3, 4)
    
    Call ga.PutAt(GNumeric(0), 0, 2)
    Call ga.PutAt(GNumeric(1), 1, 2)
    Call ga.PutAt(GNumeric(2), 2, 2)
    
    Dim Column As GenericArray
    Set Column = ga.SlizeColumn(2)
    
    Dim Element As GNumeric
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
        Set List = GenericOrderedList.AsList(GString, GNumeric, GNumeric)
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
    Dim s As GString
    Set s = GString("Ich bin ein Fuchs")
    
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
   
    Dim i As Long, N As Long
    N = 1000
        
    Set List = GenericArray.Build(N)
    ReDim x(1 To N) As IGeneric
    
    t.StartCounter
    For i = 1 To N
        Set List(i) = GNumeric.Build(i)
    Next
    Debug.Print t.TimeElapsed
    
    t.StartCounter
    For i = 1 To N
        Set x(i) = GNumeric.Build(i)
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
    
    Dim i As Long, N As Long
    N = 35
    
    Dim List As GenericOrderedList
    Set List = GenericOrderedList.Build(N)
   
    For i = 1 To N
'        Call List.Add(GenericPair(GNumeric(i), GNumeric(i)))
        Call List.Add(GNumeric(i))
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
    
    Set S1 = GString("asasasfkfnkdfndjcfv falmxxs ejf")
    Set S2 = GString("asasasfkfnkdfndjcfv falmxxs ejf")
    
    Dim i As Long, N As Long
    
    t.StartCounter
    N = 10000
    For i = 1 To N
        S1.Equals S2
    Next
    Debug.Print t.TimeElapsed
     
End Sub

Sub TestListIterator()

    Dim t As CTimer
    Set t = New CTimer
    
    Dim l As GenericOrderedList
    Set l = GenericOrderedList.Build

    Dim i As Long, N As Long
    
    N = 15
    For i = 1 To N
        Call l.Add(GenericPair(GNumeric(i), GNumeric(i)))
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

    Dim i As Long, N As Long
    
    N = 50
    For i = 1 To N
        Call sl.Add(GNumeric(i))
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
    
    Dim i As Long, N As Long, j As Long
    N = 30
    
    If RandomList Is Nothing Then
        Set RandomList = GenericOrderedList.Build
        For i = 1 To N
            Call RandomList.Add(GenericPair(GNumeric(i), GNumeric(i)))
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
    Debug.Print N & " elements added :: "; t.TimeElapsed

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

Sub TestSortedList()
    
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    
    Dim sl As GenericSortedList
    Dim i As Long, N As Long
    
    Set sl = GenericSortedList.AsList(IGenericComparer, GNumeric(7), GNumeric(9), GNumeric(3), GNumeric(1), GNumeric(10), GNumeric(11), GNumeric(13), GNumeric(6), GNumeric(8))
    Call sl.AddAll(GenericArray.AsArray(GNumeric(3), GNumeric(4), GNumeric(1), GNumeric(5), GNumeric(4), GNumeric(2), GNumeric(0), GNumeric(12), GNumeric(3)))

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
    Call tree.DoUnionWith(GenericArray.AsArray(GNumeric(3), GNumeric(4), GNumeric(1), GNumeric(5), GNumeric(2), GNumeric(0), GNumeric(3)))
    
    Call tree.DoExeceptWith(tree)
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
    
    Debug.Print C.IndexOf(GNumeric(10))
    Debug.Print C.IndexOf(GNumeric(1))
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
        Call C.Add(GString("Key: " & i), GString("Value: " & i))
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
        Call .PutAt(GString("b"), 13)
        Call .PutAt(GString("c"), 14)
        Call .PutAt(GString("a"), 15)
        Call .PutAt(GString("h"), 16)
        Call .PutAt(GString("s"), 17)
        Call .PutAt(GString("d"), 18)
        Call .PutAt(GString("zz"), 19)
        Call .PutAt(GString("c"), 20)
        Call .PutAt(GString("x"), 21)
        Call .PutAt(GString("e"), 22)
        Call .PutAt(GString("t"), 23)
        Call .PutAt(GString("a"), 24)
    
        Call .PutAt(GString("a"), 50)
        Call .PutAt(GString("c"), 51)
        Call .PutAt(GString("a"), 52)
        Call .PutAt(GString("j"), 53)
        Call .PutAt(GString("s"), 54)
        Call .PutAt(GString("ö"), 55)
        Call .PutAt(GString("q"), 56)
        Call .PutAt(GString("k"), 57)
        Call .PutAt(GString("x"), 58)
        Call .PutAt(GString("h"), 59)
        Call .PutAt(GString("t"), 60)
        Call .PutAt(GString("a"), 61)
    
        Call .PutAt(GString("z"), 70)
        Call .PutAt(GString("h"), 71)
        Call .PutAt(GString("t"), 72)
        Call .PutAt(GString("ä"), 73)
    
        Call .PutAt(GString("c"), 80)
        Call .PutAt(GString(""), 81)
        Call .PutAt(GString("e"), 82)
        Call .PutAt(GString("f"), 83)
        Call .PutAt(GString("d"), 84)
        Call .PutAt(GString("zz"), 85)
        Call .PutAt(GString("c"), 86)
        Call .PutAt(GString("x"), 87)
        Call .PutAt(GString("e"), 88)
        Call .PutAt(GString("f"), 89)
        Call .PutAt(GString("a"), 90)
        Call .PutAt(GString("a"), 99)
        
        t.StartCounter
        Call .Sort(descending, Comparer:=Nothing)
        Debug.Print t.TimeElapsed
        
        For i = .LowerBound To .Length - 1
            If Not ga.ElementAt(i) Is Nothing Then _
                Debug.Print "i: " & i & "  " & ga.ElementAt(i)
        Next

        Debug.Print .BinarySearch(GString("zzz"), descending, .LowerBound, .Length, GenericValue.Comparer)
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
    Call List.Add(GString("test0"))
    Call List.AddAll(GenericArray.AsArray(GString("test1"), GString("test3"), GString("test2"), GString("test4"), GString("test5"), GString("test3")))
    With List.Elements.Iterator
        Do While .HasNext(Item)
            Debug.Print Item
        Loop
    End With
    
    Call List.RemoveAll(GenericArray.AsArray(GString("test3"), GString("test1")))
    
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
    Set hm = GenericLinkedMap.Build(1000)
    Dim t As CTimer
    Set t = New CTimer
    
    t.StartCounter
    For i = 1 To 500
        Call hm.Add(GString("Key " & i), GNumeric(i))
    Next
    Debug.Print t.TimeElapsed
    Debug.Print "Original"
    Debug.Print System.Generic(hm)
    
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
        Call hm.Add(GString("Key " & i), GNumeric(i))
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


