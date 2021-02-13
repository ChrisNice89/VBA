Attribute VB_Name = "Modul1"

'@Folder "Entwicklung"

Option Explicit

Private sql As GenericSqlManager
Private RandomList As GenericList

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
    
    With GString.AsciiList.Sort.Iterator(119, 136)
        Do While .HasNext(Char)
            Debug.Print Char.ToString
        Loop
    End With
  
    Debug.Print t.TimeElapsed
End Sub

Sub TestCreaet()
    
    Dim Result As Boolean
    Dim t As CTimer
    Set t = New CTimer
    Dim i As Long
    Dim g As IGeneric
    
    t.StartCounter
    For i = 1 To 1000
        Set g = GBool.Build(True).Invert
'        Result = g.IsRelatedTo(g)
    Next
    Debug.Print t.TimeElapsed
    
End Sub

Public Sub TestSql()

    
'    Set sql = GenericSqlManager.BuildSqlConnection(ServerName:="192.168.2.112", InitialCatalog:="TEST", User:="SA", Password:="******")
    Set sql = GenericSqlManager.BuildAccessConnection(Path:="C:\Users\cnitz\Desktop\iCAT Neu\Backend\Vers. 2.5\2020-02-24 iCAT-Backend Vers. 2.5.accdb", Filepassword:="OpenSesame")
    
'    Call Sql.Execute(CreateTables.Test)

    Dim Updates As GenericLinkedMap
    Set Updates = GenericLinkedMap.Build
    
    Call Updates.Add(GString("Feld1"), GNumeric(100))
    Call Updates.Add(GString("Feld2"), GNumeric(101))
    Call Updates.Add(GString("Feld3"), GNumeric(102))
    Call Updates.Add(GString("Feld4"), GNumeric(103))
    Call Updates.Add(GString("Feld5"), GNumeric(104))
    Call Updates.Add(GString("Feld6"), GNumeric(105))
    
'    Call sql.UpdateWhere(GString("TEST"), GenericPair(GString("ID"), GNumeric(1)), Updates)
'    Call sql.InsertInto(GString("TEST"), Updates)


    Dim Row As GenericArray, Item As IGenericValue
   
    With sql.Query("Select * FROM tblG_00_Basis WHERE ID<5") ', GenericArray.Create(GNumeric(1), GNumeric(2)))
        Do While .HasNext(Row)
            With Row.Iterator
                Do While .HasNext(Item)
'                    Debug.Print Item.ToValue
                Loop
            End With
        Loop
    End With
    
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
    
    Call ga.SetValue(GNumeric(0), 0, 2)
    Call ga.SetValue(GNumeric(1), 1, 2)
    Call ga.SetValue(GNumeric(2), 2, 2)
    
    Dim Column As GenericArray
    Set Column = ga.SlizeColumn(2)
    
    Dim Element As GNumeric
    With Column.Iterator
        Do While .HasNext(Element)
            Debug.Print Element.Value
        Loop
    End With
    
    'Insert/ Copy Column into Matrix first column
    Call GenericArray.CopyArrays(Column, Column.LowerBound, ga, ga.LowerBound, Column.Length)
    Debug.Print ga.GetValue(0, 0).Equals(ga.GetValue(0, 2))
    
    Call ga.Transpose
    Debug.Print ga.GetValue(2, 0).Equals(Column(0))
     Call Column.Clear
    
End Sub
Sub testArrayConstructor()
    
    Dim t As CTimer
    Set t = New CTimer
    
    Dim List As GenericList
    Dim i As Long
     
    t.StartCounter
   
    For i = 1 To 100
        Set List = GenericList.Create(GString, GNumeric, GNumeric)
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
    
    With s.ToArray.Shuffle(12, 5).Iterator(12, 5)
        Do While .HasNext(Char)
            Debug.Print Char
        Loop
    End With
    
    With s.Split(" ").Iterator
        Do While .HasNext(Char)
            Debug.Print Char
        Loop
    End With
    
    Debug.Print s.ElementAt(1).Contains("i")
    
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
    
    Dim List As GenericList
    Set List = GenericList.Build(N)
   
    For i = 1 To N
'        Call List.Add(GenericPair(GNumeric(i), GNumeric(i)))
        Call List.Add(GNumeric(i))
    Next
    Dim Item As IGeneric
    
    With List.Shuffle.Iterator(5, 10)
        Do While .HasNext(Item)
            Debug.Print Item
        Loop
    End With
    
    t.StartCounter
    
    With List.Sort(ascending).Iterator(5, 10)
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

Sub TestOrderedMap()
    
    Dim t As CTimer
    Set t = New CTimer
    
    Dim Map As GenericOrderedMap
    Set Map = GenericOrderedMap.Build
    
    Dim Imap As IGenericDictionary
    Set Imap = GenericLinkedMap.Build
    
    Dim i As Long, N As Long
    
    N = 30
    t.StartCounter
    For i = 0 To N - 1
        Call Imap.Add(GNumeric(i), GNumeric(i))
    Next
    
    t.StartCounter
    Call Map.AddAll(Imap)
    Debug.Print t.TimeElapsed
    
    Dim c As GenericOrderedMap
    Set c = System.Clone(Map)
    
    Dim Item As IGeneric
    t.StartCounter
    With c.GetKeys.Iterator
        Do While .HasNext(Item)
            Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed
    
End Sub

Sub TestListIterator()

    Dim t As CTimer
    Set t = New CTimer
    
    Dim l As GenericList
    Set l = GenericList.Build

    Dim i As Long, N As Long
    
    N = 5000
    For i = 1 To N
        Call l.Add(GenericPair(GNumeric(i), GNumeric(i)))
    Next
    
    Dim c As GenericList
    Set c = System.Clone(l)
    
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

    Dim i As Long, N As Long
    
    N = 50
    For i = 1 To N
        Call sl.Add(GNumeric(i))
    Next
    
    Dim c As GenericSortedList
    Set c = System.Clone(sl)
    
    t.StartCounter
    Dim Item As IGeneric
    With c.Iterator
        Do While .HasNext(Item)
           Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed

End Sub

Sub TestMaps()
    
    Dim t As CTimer
    
    Dim Map As IGenericDictionary
    Set Map = GenericSortedList.Build() 'GenericSortedList.Build()'enericOrderedMap.Build
    
    Dim i As Long, N As Long, j As Long
    N = 30
    
    If RandomList Is Nothing Then
        Set RandomList = GenericList.Build
        For i = 1 To N
            Call RandomList.Add(GenericPair(GNumeric(i), GNumeric(i)))
        Next
        Call RandomList.Shuffle
        Call RandomList.Sort(ascending, GenericPair)
    End If
 
    Dim p As GenericPair
    Dim Item As IGeneric
    
    Set t = New CTimer
    t.StartCounter
    For i = RandomList.First To RandomList.Count - 1
        Set p = RandomList(i)
        Call Map.Add(p.Key, p.Value)
    Next
    Debug.Print N & " :: "; t.TimeElapsed

    For i = RandomList.First To RandomList.Count - 1
        Set p = RandomList(i)
        Set Item = Map.Item(p.Key)
    Next
  
    Dim ga As GenericArray
    Set ga = GenericArray.Build(Map.Count)
   
    
    Dim GenericPairComparer As IGenericComparer
    Set GenericPairComparer = GenericPair

    For i = ga.LowerBound To ga.Length - 1
        Debug.Print ga.ElementAt(i)
    Next
'
    t.StartCounter
    With Map.Iterator()
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
    
    Set sl = GenericSortedList.Create(IGenericComparer, GNumeric(3), GNumeric(4), GNumeric(1), GNumeric(5), GNumeric(2), GNumeric(0))
    Debug.Print t.TimeElapsed
    
    Dim c As GenericSortedList
    Set c = System.Clone(sl)
    
    Dim Item As IGeneric
    
    With c.Iterator '(0, 3)
        Do While .HasNext(Item)
           Debug.Print Item
        Loop
    End With
    Debug.Print t.TimeElapsed

End Sub

Sub TestTree()
    
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    
    Dim tree As GenericSortedSet
    Set tree = GenericSortedSet.Build
    
    Dim i As Long
    
    For i = 30000 To 1 Step -1
        Call tree.Add(GNumeric(i))
    Next
    Debug.Print t.TimeElapsed

'    Dim N As IGeneric
'    Set N = tree.ElementAt(1)
'    Debug.Print N.ToString
    Dim c As IGenericReadOnlyList
    Set c = System.Clone(tree)
    Call tree.Clear
    Set tree = Nothing
    
    Debug.Print c.IndexOf(GNumeric(10))
    Debug.Print c.IndexOf(GNumeric(1))
    Dim Item As IGeneric
    
    t.StartCounter
    
    With c.Iterator
        Do While .HasNext(Item)
            Debug.Print Item.ToString
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
    
    Call List.AddAll(c.Iterator()) 'size is unknown
    'Call List.AddAll(c)' faster because size is known
   
    Dim Clone As IGenericReadOnlyList
    Set Clone = System.Clone(List.AsReadOnly)
        
    Dim ga As GenericArray
    Set ga = GenericArray.Build(Clone.Count)
    
    Call Clone.CopyTo(ga, ga.LowerBound)
    Call System.Dispose(Clone)
       
    For i = ga.LowerBound To ga.Length
        Debug.Print ga(i)
    Next

    Debug.Print ga.IndexOf(List(10))

End Sub

Sub TestArray2()
    
    Dim i As Long
    Dim ga As GenericArray
    Set ga = GenericArray.Build(100)
    Dim t As CTimer
    Set t = New CTimer
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
        Call .SetValue(GString("a"), 99)
        t.StartCounter
        Call .Sort(descending)
        Debug.Print t.TimeElapsed
        
        For i = 1 To .Length - 1
            If Not ga(i) Is Nothing Then _
                Debug.Print "i: " & i & "  " & ga(i)
        Next

        Debug.Print .BinarySearch(GString("zzz"), 1, .Length - 1, descending, IGenericValue.Comparer)
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
    Set List2 = System.Clone(List)
    Debug.Print List2.Count
    Call List.Insert(500, GString("eingefügt an 500"))
    Debug.Print List(500)
    Debug.Print List2(500)

    Dim List3 As GenericList
    Set List3 = List.GetRange(500, 10)
    
    Dim readOnly As IGenericReadOnlyList
    Set readOnly = List3.AsReadOnly
    Debug.Print readOnly(1)
    Debug.Print readOnly(9)
    
    Debug.Print readOnly.Count
    Set List = Nothing

End Sub

Sub testMap()

    Dim i As Long, Pair As GenericPair, Item As IGeneric
    
    Dim hm As GenericLinkedMap
    Set hm = GenericLinkedMap.Build()
    Dim t As CTimer
    Set t = New CTimer
    
    t.StartCounter
    For i = 1 To 10000
        Call hm.Add(GString("Key " & i), GNumeric(i))
    Next
    
    With hm.Iterator
        Do While .HasNext(Pair)
'            Debug.Print Pair.Key
        Loop
    End With
    
    With hm.GetValues.Iterator(0, 10)
        Do While .HasNext(Item)
'            Debug.Print Item.ToString
        Loop
    End With
    
    Set hm = Nothing
    Debug.Print t.TimeElapsed
    
End Sub

