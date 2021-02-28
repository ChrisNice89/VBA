Attribute VB_Name = "Modul1"

'@Folder "Entwicklung"

Option Explicit

Private Sql As GenericSqlManager
Private RandomList As GenericOrderedList

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

Public Sub TestSql()

    
    Set Sql = GenericSqlManager.BuildSqlConnection(ServerName:="192.168.0.186", InitialCatalog:="TEST", User:="TestUser", Password:="OpenSesame")
'    Set Sql = GenericSqlManager.BuildAccessConnection(Path:="C:\Users\cnitz\Desktop\iCAT Neu\Backend\Vers. 2.5\2020-02-24 iCAT-Backend Vers. 2.5.accdb", Filepassword:="OpenSesame")
    
'    Call Sql.Execute(CreateTables.Test)

    Dim Updates As GenericLinkedMap
    Set Updates = GenericLinkedMap.Build
    
'    Call Updates.Add(GString("KNE"), GNumeric(100))

'    Call Sql.UpdateWhere(GString("TEST"), Updates, GenericPair(GString("ID"), GNumeric(1)))
'    Call Sql.InsertInto(GString("TEST"), Updates)


    Dim Row As GenericOrderedMap, Item As IGeneric
   
    With Sql.SelectWhere(GString("TEST"), GenericArray.AsArray(GString("KNE")), GenericPair(GString("KNE"), GNumeric(100)))
        Do While .HasNext(Row)
            With Row.Elements.Iterator
                Do While .HasNext(Item)
                    Debug.Print Item
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
    With Column.Elements.Iterator
        Do While .HasNext(Element)
            Debug.Print Element.Value
        Loop
    End With
    
    'Insert/ Copy Column into Matrix first column
    Call GenericArray.CopyArrays(Column, Column.LowerBound, ga, ga.LowerBound, Column.Length)
    Debug.Print ga.GetValue(0, 0).Equals(ga.GetValue(0, 2))
    
    Call ga.Transpose
    Debug.Print ga.GetValue(2, 0).Equals(Column(0))
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

Sub TestOrderedMap()
    
    Dim t As CTimer
    Set t = New CTimer
    
    Dim Map As GenericOrderedMap
    Set Map = GenericOrderedMap.Build
    
    Dim Imap As GenericLinkedMap
    Set Imap = GenericLinkedMap.Build
    
    Dim i As Long, N As Long
    
    N = 30
    t.StartCounter
    For i = 1 To N
        Call Imap.Add(GNumeric(i), GNumeric(i))
    Next
    
    t.StartCounter
    Call Map.AddAll(Imap)
    Debug.Print t.TimeElapsed
    
    Dim c As GenericOrderedMap
    Set c = System.Clone(Map)
    
    Dim Item As IGeneric
    t.StartCounter
    With c.GetKeys.Elements.Iterator()
        Do While .HasNext(Item)
            Debug.Print Item
        Loop
    End With
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

    Dim i As Long, N As Long
    
    N = 50
    For i = 1 To N
        Call sl.Add(GNumeric(i))
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
    
    Set sl = GenericSortedList.Build(Comparer:=IGenericComparer)
    Call sl.AddAll(GenericArray.AsArray(GNumeric(3), GNumeric(4), GNumeric(1), GNumeric(5), GNumeric(2), GNumeric(0), GNumeric(3)))

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
    
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    
    Dim tree As GenericSortedSet
    Set tree = GenericSortedSet.Build
    
    Dim i As Long
    
    For i = 5 To 1 Step -1
        Call tree.Add(GNumeric(i))
    Next
    
    For i = 0 To 4
        Debug.Print tree.ElementAt(i)
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
    Dim c As IGenericReadOnlyList
    Set c = System.Clone(tree)
    Call tree.Elements.Clear
    Set tree = Nothing
    
    Debug.Print c.IndexOf(GNumeric(10))
    Debug.Print c.IndexOf(GNumeric(1))
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
    
    Dim i As Long
    For i = 1 To 10
        Call c.Add(GString("Key: " & i), GString("Value: " & i))
    Next

    Dim List As GenericOrderedList
    Set List = GenericOrderedList.Build
    
    Call List.AddAll(c.Elements.Iterator()) 'size is unknown
    'Call List.AddAll(c)' faster because size is known
   
    Dim Clone As IGenericReadOnlyList
    Set Clone = List.Elements.Copy
        
    Dim ga As GenericArray
    Set ga = GenericArray.Build(Clone.Elements.Count)
    
    Call Clone.Elements.CopyTo(ga, ga.LowerBound)
       
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
        
        For i = .LowerBound To .Length - 1
            If Not ga.ElementAt(i) Is Nothing Then _
                Debug.Print "i: " & i & "  " & ga.ElementAt(i)
        Next

        Debug.Print .BinarySearch(GString("zzz"), descending, .LowerBound, .Length, IGenericValue.Comparer)
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
'            Debug.Print item
        Loop
    End With
    
    Call List.RemoveAll(GenericArray.AsArray(GString("test3"), GString("test1")))
    Dim ArrayList As GenericArray
    Set ArrayList = GenericArray.Build(List.Elements.Count)
    Call List.CopyTo(ArrayList, ArrayList.LowerBound)
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
    Set hm = GenericLinkedMap.Build(10)
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

