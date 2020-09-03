Attribute VB_Name = "Modul1"



'@Folder "Entwicklung"

Option Explicit


Sub TestArrayiterator()


End Sub

Sub testEnumerable()
    Call Skynet.TestString
    Call Skynet.TestInteger
    Call Skynet.TestContains
'

'    Call Skynet.TestInteger2(15000)

'    Call Skynet.TestRange
End Sub
Sub TestTree2()
    
    Dim tree As GenericTree
    Set tree = GenericTree.Build
    
    Dim i As Long, n As Long
    Dim t As Timer
    Set t = New Timer
    t.StartCounter
    n = 50000
    For i = n To 1 Step -1
        Call tree.Add(GNumeric(i), GNumeric(i))
    Next
    Debug.Print t.TimeElapsed
    Debug.Print tree.GetMax
    Dim ga As GenericArray

    Set ga = tree.GetKeys(True)
'    For i = ga.LowerBound To ga.Length
'        Debug.Print ga(i)
'    Next
''
'
'    With tree.Enumerator
'        Do While .MoveNext
'            Debug.Print .Current
'        Loop
'    End With

End Sub

Sub TestTree()
    
    Dim tree As GenericTree
    Set tree = GenericTree.Build

    Call tree.Add(GString("c"), GString("c"))
    Call tree.Add(GString("z"), GString("z"))
    Call tree.Add(GString("a"), GString("a"))
    Call tree.Add(GString("v"), GString("v"))
    Call tree.Add(GString("f"), GString("f"))
    Call tree.Add(GString("l"), GString("l"))
    Call tree.Add(GString("d"), GString("d"))
    Call tree.Add(GString("b"), GString("b"))
    Call tree.Add(GString("e"), GString("e"))
    Call tree.Add(GString("x"), GString("x"))
    Call tree.Add(GString("h"), GString("h"))
    Call tree.Add(GString("g"), GString("g"))
    Call tree.Add(GString("zz"), GString("zz"))
    
    Debug.Print tree.GetMin
        
    Dim e As IGenericEnumerator
    Set e = tree
    
    Do While e.MoveNext
        Debug.Print e.Current
    Loop


    Dim ga As GenericArray
    
    Set ga = tree.GetKeys(True)
    
    Dim i As Long
    For i = 1 To ga.Length
        Debug.Print ga(i)
    Next
    
    Debug.Print ga.IndexOf(GString("zz"))
   
End Sub
Sub TestSortedList()

    Dim sl As GenericSortedList
    Set sl = GenericSortedList.Build
    
    Dim i As Long
    
    Call sl.Add(GString("c"), GString("c"))
    Call sl.Add(GString("z"), GString("z"))
    Call sl.Add(GString("a"), GString("a"))
    Call sl.Add(GString("f"), GString("f"))
    Call sl.Add(GString("aa"), GString("aa"))
    Call sl.Add(GString("d"), GString("d"))
    Call sl.Add(GString("b"), GString("b"))
    Call sl.Add(GString("e"), GString("e"))
    
    For i = 1 To sl.Count
        If Not sl.ElementAt(i) Is Nothing Then _
            Debug.Print i & " / " & sl.ElementAt(i)
    Next
    
End Sub
Sub TestArrayEnumerator()

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

    Dim list As GenericList
    Set list = GenericList.Build
    
    Call list.AddAll(C.GetKeys)
   
    Dim Clone As IGenericReadOnlyList
    Set Clone = Skynet.Clone(list.AsReadOnly)
        
    Dim ga As GenericArray
    Set ga = GenericArray.Build(Clone.Elements.Count)
    
    Call Clone.Elements.CopyTo(ga, ga.LowerBound)
    Call Skynet.Dispose(Clone)
       
    For i = ga.LowerBound To ga.Length
        Debug.Print ga(i)
    Next

    Debug.Print ga.IndexOf(list(10))
    
    Call C.Remove(C.GetKeys(9))
    
End Sub
Sub TestGenericPairXY()
    
    Dim C As New VBA.Collection
        
    Dim pair1 As IGeneric
    Set pair1 = GenericPairXY(GString("A"), Nothing)
    
    
    Dim pair2 As IGeneric
    Set pair2 = GenericPairXY(GString("A"), C)
    
    Debug.Print pair1.Equals(pair2)
End Sub

Sub asadsfsd()

Dim x(1 To 10) As IGeneric
Debug.Print Skynet.BinarySearch(x, Nothing, 1, 1, Ascending)


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
    Call ga.Sort(Descending)

    For i = 1 To ga.Length
        If Not ga(i) Is Nothing Then _
            Debug.Print "i: " & i & "  " & ga(i)
    Next
    
    Debug.Print ga.BinarySearch(GString("a"), 1, ga.Length, Descending)
    Call ga.Reverse

End Sub

Public Sub Redim1()
    
    Dim i As Long
    ReDim a(1 To 50000) As IGeneric
    
    For i = 1 To 50000
        Set a(i) = GString("test" & i)
    Next
    Dim t As Timer
    Set t = New Timer
    t.StartCounter
    ReDim Preserve a(1 To 100000)
    Debug.Print t.TimeElapsed

End Sub

Public Sub Redim2()
    Dim t As Timer
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

    Dim list As GenericList
    Set list = GenericList.Build()
    For i = 1 To 5
        list.Add GString("testX" & i)
    Next

    Call list.InsertAll(3, a)
    
    Call list.Reverse
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
    
    Dim t As Timer
    Set t = New Timer
    
    Dim a(1 To 10000) As IGeneric
    Set ga = GenericArray.Build(10000)
    
    t.StartCounter
    For i = 1 To ga.Length
        Set ga(i) = GString("Test" & i)
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
    
    
    Dim readOnly As IGenericReadOnlyList
    Set readOnly = ga.AsReadOnly
    
End Sub

Public Sub ListTest()

    Dim i As Long
    Dim list As GenericList
    Set list = GenericList.Build()
    
    Dim t As Timer
    Set t = New Timer
    t.StartCounter
    For i = 1 To 10000
        Call list.Add(GString("test" & i))
    Next
    Debug.Print t.TimeElapsed
'
'    Debug.Print List.IndexOf(GString("test" & 999), 1, 999)
'    Call List.Insert(500, GString("eingefügt an 500"))
'    Debug.Print List.IndexOf(GString("eingefügt an 500"))
'
    Dim List2 As GenericList
    Set List2 = Skynet.Clone(list)
    Debug.Print List2.Count
    Call list.Insert(500, GString("eingefügt an 500"))
    Debug.Print list(500)
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
    Dim k As GString
    Dim v As GString
    
    Dim hm As GenericMap
    Set hm = GenericMap.Build()
    
    For i = 1 To 20000
        Call hm.Add(GString("Key" & i), GString("Value" & i)) 'TString("Value" & i)
    Next
    Debug.Print hm(GString("Key1"))
    Dim hm2 As GenericMap
    Set hm2 = Skynet.Clone(hm)
    
    Debug.Print hm2 Is hm
    Debug.Print Skynet.Generic(hm2).Equals(hm)
    Debug.Print hm2.Item(GString("Key" & 1)).ToString
    
    Call hm.Add(GString("Key" & 1), GString("ValueNew" & 1))
    Debug.Print hm2.Item(GString("Key" & 1)).ToString
    Debug.Print hm.Item(GString("Key" & 1)).ToString
    
    Dim Keys As GenericArray
    Set Keys = hm.GetKeys
    Call Keys.Sort(Ascending)
    
    i = Keys.IndexOf(GString("Key" & 5))
    Debug.Print Keys(i)
 
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


