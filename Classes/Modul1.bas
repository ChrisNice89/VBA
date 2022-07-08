Attribute VB_Name = "Modul1"
Option Explicit
'@Ignore MoveFieldCloserToUsage
Private RandomList As GenericOrderedList

Private Const x7FFFFFFF As Long = &H7FFFFFFF
Private Const x80000000 As Long = &H80000000

Private Type SomeStruct
    Length As Long
End Type

Private Declare PtrSafe Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As LongPtr) As Long 'Works
Private Declare PtrSafe Function SysAllocStringLen Lib "oleaut32" (ByVal psz As LongPtr, ByVal cLen As Long) As String
Private Declare PtrSafe Function SysAllocString Lib "oleaut32.dll" (ByVal pBSTR As LongPtr) As String

Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal psz As Long, ByVal cbLen As Long) As String

Private Declare Function SysStringLen Lib "oleaut32" (ByVal psz As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As LongPtr) As Long
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As String) As Long
Private Declare Function lstrcopyA Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Public Sub TestSort()
    
    Dim List As GenericOrderedList
    Set List = GenericOrderedList.Build
    
    Dim i As Long
    
    For i = 35 To 1 Step -1
        Call List.Add(gInt.Of(i))
    Next
    For i = 35 To 1 Step -1
        Call List.Add(gInt.Of(i))
    Next
    For i = 35 To 1 Step -1
        Call List.Add(gInt.Of(i))
    Next
    For i = 35 To 1 Step -1
        Call List.Add(Nothing)
    Next
    Call List.Shuffle.Shuffle

    For i = 1 To 15
        Call List.Add(Nothing)
    Next
'
    Call List.Shuffle
    
    With List.Sort(Ascending, IGenericValue).Range
        Do While .HasNext()
            If Not .Current Is Nothing Then
                Debug.Print .Current
            Else
                Debug.Print "nothing"
            End If
        Loop
    End With
    
   
End Sub

Private Sub Inject(ParamArray Elements() As Variant)
    
    Dim i As Long
    For i = 0 To UBound(Elements)
        Debug.Print Elements(i)
    Next
    
End Sub
Public Sub TestContainsAll()
    
    Dim A As GenericOrderedList
    Set A = GenericOrderedList.Build
    
    Dim B As GenericSortedList
    Set B = GenericSortedList.Build
    
    Dim i As Long
    
    For i = 1 To 10000
        Call A.Add(gString.Of("Key: " & i))
    Next
   
    Call B.AddAll(A.Shuffle)
    
    Debug.Print B.ContainsAll(A)
    Debug.Print A.Stream.ContainsAll(B, Nothing, IGenericValue)
    
End Sub

Public Function StrFromAnsiPtr(ByVal lpStr As Long) As String
    StrFromAnsiPtr = SysAllocStringLen(lpStr, 1) 'SysAllocStringByteLen(lpStr, lstrlenW(lpStr) * 6)
End Function

Sub TestMapVsSet()
    
    Dim t As CTimer
    Set t = New CTimer
    
    Dim List As GenericOrderedList
    Set List = GenericOrderedList.Build
    
    Dim HashSet As GenericHashSet
    Set HashSet = GenericHashSet.Build(500000)
    
    Dim i As Long, s As IGenericValue
    
    For i = 1 To 20000
        Set s = gString.Of("Key: " & i)
        Call List.Add(s)
    Next
    
    Call List.Shuffle
    
'    t.StartCounter
'    For i = List.First To List.Last
'        If Not Map.TryAdd(List.GetAt(i), Nothing) Then
'            Debug.Print "hier"
'        End If
'    Next
'    Debug.Print t.TimeElapsed
'
    t.StartCounter
    For i = List.First To List.Last
     If Not HashSet.TryAdd(List.GetAt(i)) Then
            Debug.Print "hier"
        End If
    Next
    Debug.Print t.TimeElapsed

    
End Sub

Sub TestUnorderedSet()
    
    Dim List As GenericOrderedList
    Set List = GenericOrderedList.Build
    
    Dim h1 As GenericHashSet
    Set h1 = GenericHashSet.Build
  
    Dim i As Long, s As IGenericValue
    
    For i = 1 To 10000
        Set s = gString.Of("Key: " & i)
        Call List.Add(s)
    Next
    
    Call List.Shuffle
    
    For i = List.First To List.Last
        If Not h1.TryAdd(List.GetAt(i)) Then
            Debug.Print "hier"
        End If
    Next
   
    Call h1.DoExcept(h1.Elements.Copy)
    
'    Call List.Shuffle
'
'
'     For i = List.First To List.Last
'        If Not h1.Contains(List.GetAt(i)) Then
'            Debug.Print "hier"
'        End If
'    Next
'
'    Debug.Print h1.EqualsTo(h1.Elements.Copy)
    
'    With h1.Elements.Iterator
'
'        Do While .HasNext
'            Debug.Print .Current.ToString
'        Loop
'
'    End With

End Sub
Sub StringTest()
    
    
    Dim x As String
    x = "ä"
    
    Dim Test As gString
    Set Test = gString.Of(x)
    
    Debug.Print Test.StartsWith("Ä", False)
    
    Debug.Print gString.Of("  abc   ").Trim.ToString
End Sub

'Returns a copy of a null-terminated Unicode string (LPWSTR/LPCWSTR) from the given pointer
'@Ignore NonReturningFunction
Public Function GetStrFromPtrW(ByVal Ptr As LongPtr) As String
    SysReAllocString VarPtr(GetStrFromPtrW), Ptr
End Function

Sub TestSysReAllocString()

  
    '@Ignore VariableNotAssigned
    Dim y As String
    
    Dim t As CTimer
    Set t = New CTimer
    
    Dim i As Long
    t.StartCounter
    For i = 1 To 10000
       Call SysReAllocString(VarPtr(y), StrPtr("new"))
    Next
    Debug.Print t.TimeElapsed
    Debug.Print y
'
End Sub


Sub TestIntersection()

    Dim s1 As GenericHashSet
    Dim s2 As GenericHashSet
    
    Dim List As GenericOrderedList
    Set List = GenericOrderedList.Build()
    Dim i As Long
    For i = 1 To 600
        Call List.Add(gString.Of("Key: " & i))
    Next
    
    Set s1 = GenericHashSet.Build().DoUnion(List.Shuffle).TrimExcess
    
    Call List.RemoveAll(GenericArray.Of(gString.Of("Key: " & 1), gString.Of("Key: " & 2), gString.Of("Key: " & 3)))

    With s1.SymmetricDifference(GenericHashSet.BuildFrom(List))
        With .Elements.Iterator
            Do While .HasNext
                Debug.Print .Current.Instance.ToString
            Loop
        End With
    End With
'
    With s1.SymmetricDifference(List).Elements.Iterator
        Do While .HasNext
            Debug.Print .Current.Instance.ToString
        Loop
    End With
'
    With s1.Difference(GenericHashSet.BuildFrom(List)).Elements.Iterator
        Do While .HasNext
            Debug.Print .Current.ToString
        Loop
    End With
    
    With s1.DoExcept(GenericHashSet.BuildFrom(List)).Elements.Iterator
        Do While .HasNext
            Debug.Print .Current.ToString
        Loop
    End With
'
    
    Set s2 = s1.Elements.Copy
    
    Debug.Print s1.Equals(s2)
    
End Sub

Sub testSortedSet()

    Dim s1 As GenericSortedSet
    Dim s2 As GenericSortedSet
    
    Dim List As GenericOrderedList
    Set List = GenericOrderedList.Build()
    Dim i As Long
    For i = 1 To 50
        Call List.Add(gInt.Of(i))
    Next
    List.Shuffle
    Dim t As CTimer
    Set t = New CTimer
    
    t.StartCounter
    Set s1 = GenericSortedSet.Build().DoUnion(List)
    Debug.Print t.TimeElapsed
    
'    t.StartCounter
'    For i = List.First To List.Last
'
'        If Not S1.GetAt(i).ToLong = i + 1 Then
'            Debug.Print "hier"
'        End If
'    Next
'    Debug.Print t.TimeElapsed
    
    t.StartCounter
    With s1.GetBetween(gInt.Of(20), gInt.Of(35)).Elements.Iterator
        Do While .HasNext
            Debug.Print .Current
        Loop
    End With
    Debug.Print t.TimeElapsed
    
    
    Debug.Print s1.GetMin.ToLong
    Debug.Print s1.GetMax.ToLong
    
 
    Set s2 = GenericSortedSet.Build().DoUnion(List.Shuffle)
   
    Dim Comparer As IGenericComparer
    Set Comparer = GenericSortedSet
    
    Debug.Print Comparer.Equality(s1, s2)
   
    Call s2.Elements.ToArray
    
    With s2.Elements.Iterator
        Do While .HasNext
            Debug.Print .Current
        Loop
    End With

End Sub

Sub TestQuery()

    Dim i As Long
    Dim Querys As GenericOrderedList
    Dim SqlManager As GenericSqlManager
    
    Dim t As CTimer
    Set t = New CTimer
    t.StartCounter
    
    Set SqlManager = GenericSqlManager.BuildAccessConnection("")
    Set Querys = GenericOrderedList.Build
    
    For i = 1 To 10000
        Call Querys.Add(SqlManager.Query(gString.Of("SELECT * FROM TEST"), Prepared).Add(gInt.Of(i)).Add(gInt.Of(i)).Add(gInt.Of(i)).Add(gInt.Of(i)))
    Next
    
    Debug.Print t.TimeElapsed
    Set Querys = Nothing
    
End Sub

Sub SliceSet()

    Dim HashSet As GenericHashSet
    Set HashSet = GenericHashSet.Build

    Dim i As Long
    For i = 1 To 100
        Call HashSet.TryAdd(gInt.Of(i))
    Next
    
    With HashSet.Slice(5, 10).Elements.Iterator
        Do While .HasNext
            Debug.Print .Current.Instance.ToString
        Loop
    End With

End Sub

Sub TestConcat()
    
    '@Ignore VariableNotAssigned
    Dim Element As Object
    Dim A As GenericOrderedList
    Set A = GenericOrderedList.Build()
    
    Dim B As GenericOrderedList
    Set B = GenericOrderedList.Build()
    
    Dim c As GenericSortedList
    Set c = GenericSortedList.Build()
    
    Dim i As Long
    For i = 1 To 10000
        Call A.Add(gInt.Of(i))
    Next
    
    For i = 10001 To 20000
        Call B.Add(gInt.Of(i))
    Next
    
    For i = 20001 To 30000
        Call c.Add(gInt.Of(i))
    Next
    i = 1
'
'    Debug.Print GenericSequence.Stream(a).Concat(b).GetAt(19).Instance.ToString
'    i = GenericSequence.Stream(a).Concat(b).Count
'
    With GenericSequence.Stream(A.Shuffle).Append(B.Shuffle).Append(c).Ascending.Distinct.ToList.Elements.Iterator
        Do While .HasNext(Element)
            Debug.Print "Element : " & i & " :: " & Element.Instance.ToString
            i = i + 1
        Loop
    End With
   
'   Debug.Print GenericSequence.Stream(List).Descending.GetAt(0).ToString
   
End Sub

Sub TestStringJoin()

    Dim List As GenericOrderedList
    Set List = GenericOrderedList.Build()
   
    Dim i As Long
    For i = 1 To 10
        Call List.Add(gString.Of("Nummer: " & i))
    Next

    Debug.Print gString.Join(List, vbNewLine).ToString
    
End Sub

Sub TestMap()

    Dim Map As GenericHashMap
    Set Map = GenericHashMap.Build()

    Dim List As GenericOrderedList
    Set List = GenericOrderedList.Build()
   
    Dim i As Long
    For i = 1 To 50000
        Call List.Add(gString.Of(String(1000, VBA.Chr$(255)) & i))
    Next
    
    Call Map.AddAll(List.Shuffle)
'
'    For i = 15000 To 0 Step -1
'        Call map.RemoveAt(i)
'        Call List.RemoveAt(i)
'    Next
    
'    With map.Range
'        Do While .HasNext
'            Debug.Print .Current.Instance.ToString
'        Loop
'    End With
    
    For i = List.First To List.Last

        If Not Map.Contains(List.GetAt(i)) Then
            Debug.Print List.GetAt(i)
        End If
    Next
    Debug.Print Map.MeanSearch
End Sub

Sub TestSetEquals()
   
   Dim t As CTimer
   Set t = New CTimer
   
   Dim s1 As GenericHashSet
   Set s1 = GenericHashSet.Build(100)
   
   Dim s2 As GenericHashSet
   Set s2 = GenericHashSet.Build(100)
   
   Dim List As GenericOrderedList
   Set List = GenericOrderedList.Build
   
    Dim i As Long
    For i = 1 To 30000
        Call List.Add(gString.Of("Key: " & i))
    Next

    Call s1.DoUnion(List.Shuffle)
    Call s2.DoUnion(List.Shuffle)
    Debug.Print s1.Equals(s2)
    
    Call s1.TryAdd(gString.Of("Key: " & 0))
    Call s1.TryRemove(gString.Of("Key: " & 1))
    Call s2.TryRemove(gString.Of("Key: " & 2))
    Call s2.TryRemove(gString.Of("Key: " & 3))
    
    t.StartCounter
    Call s1.DoSymmetricExcept(s2.Elements.ToArray)
    Debug.Print t.TimeElapsed
    
'    With s1.Elements.Iterator
'        Do While .HasNext
'            Debug.Print .Current.ToString
'        Loop
'    End With
   
End Sub

Sub testIterator()

    Dim List As GenericOrderedList
    
    Set List = GenericOrderedList.Build()
    Dim i As Long
  
    For i = 0 To 128
        Call List.Add(gInt.Of(i))
    Next
    
    Call List.Insert(5, gString.Of("Test"))
    
'    For i = List.First To List.Last
'        If List.GetAt(i) Is Nothing Then
'            Debug.Print "nothing"
'        Else
'            Debug.Print List.GetAt(i)
'        End If
'    Next
    
    With List.Elements.Iterator

        Do While .HasNext
            If .Current Is Nothing Then
                Debug.Print "nothing"
            Else
                Debug.Print .Current
            End If
        Loop
'
    End With
    
'    Call List.Elements.Clear
    
End Sub

Sub TestMulti2()
    
    Dim A As GenericArray, B As IGenericCollection
    
    Set A = GenericArray.Build(3, 3)
    
    With A
        Call .PutAt(gInt.Of(0), 0, 0)
        Call .PutAt(gInt.Of(1), 1, 0)
        Call .PutAt(gInt.Of(2), 2, 0)
        Call .PutAt(gInt.Of(5), 1, 1)
        
        Debug.Print .Elements.Contains(gInt.Of(5))
        
    End With
    
    Set B = A.Slice(0)
   
    With B.Iterator
        Do While .HasNext
            
            If Not .Current Is Nothing Then
                Debug.Print .Current.Instance.ToString
            End If
            
        Loop
    
    End With
  
End Sub

Sub TestMulti()
    
    Dim A As GenericArray, B As GenericArray
    
    Set A = GenericArray.Build(3, 3, 3)
    
    With A
        Call .PutAt(gInt.Of(0), 0, 0, 0)
        Call .PutAt(gInt.Of(1), 1, 0, 0)
        Call .PutAt(gInt.Of(2), 2, 0, 0)
        Call .PutAt(gInt.Of(5), 0, 0, 1)
        
        Call .PutAt(gInt.Of(10), 2, 2, 2)
        Debug.Print .Elements.Contains(gInt.Of(10))
        
    End With
    
    Set B = GenericArray.Build(3)
    Call A.CopyTo(B, 0, 0, 3)
    
    Call B.CopyTo(A, 0, 0, 3)
    
    With A.Elements.Iterator
        Do While .HasNext
            
            If Not .Current Is Nothing Then
                Debug.Print .Current.Instance.ToString
            End If
            
        Loop
    
    End With
  
End Sub

Sub TestEnum()

    With GenericSequence.Of(gString.Of("Christoph"), gString.Of("Nitz"), gString.Of("51")).Ascending.Source.Iterator
        Do While .HasNext
            Debug.Print .Current.ToString
        Loop
    End With
    
End Sub

Sub TestPartialSort()

    Dim i As Long
    If RandomList Is Nothing Then
        Set RandomList = GenericOrderedList.Build
        For i = 1 To 100
            Call RandomList.Add(gInt.Of(i))
        Next
        
         Call RandomList.Shuffle
    End If
    
    Dim List As GenericArray

    Set List = RandomList.Elements.ToArray
    
'TODO Is this comment still valid? =>     Debug.Print RandomList.SelectKth(17).ToString
    Call List.SortPartial(91, 500, Descending, Nothing)
    
    '@Ignore VariableNotAssigned
    Dim Element As IGenericValue
    
    With List.Elements.Iterator
        Do While .HasNext(Element)
            Debug.Print Element.GetValue
        Loop
    End With
End Sub

Sub TestStrungConcat()

    Debug.Print gString.Join(GenericSequence.Stream(GenericArray.Of(gString.Of("Christoph"), gString.Of("Nitz"), gString.Of("51"))), " ").ToString
    
End Sub

Sub TestInsertString()

    Debug.Print gString.Of("BC").Insert(2, "A").ToString

End Sub

Sub TestList()

    Dim List As GenericOrderedList
    Set List = GenericOrderedList.Build
    
    With List
        Dim i As Long
        For i = 1 To 15
            Call .Add(gInt.Of(i))
        Next
        For i = 1 To 15
            Call .Add(gInt.Of(i))
        Next
        For i = 1 To 30
            Call .Add(Nothing)
        Next
        Call .Shuffle
        Call .Sort
        
'        Set Clone = .Elements.ToArray
                
    End With
    
'    Set Sorter = Clone.Elements.Copy
'
'    Call Sorter.SortWith(Clone, Descending)
    
    With List.Elements.Iterator
        Do While .HasNext
            If .Current Is Nothing Then
                Debug.Print "Nothing"
            Else
                Debug.Print .Current.Instance.ToString
            End If
        Loop
    End With
    
End Sub

Sub TestSortedList()

    Dim List As GenericOrderedList
    Set List = GenericOrderedList.Build
    
    With List
        Dim i As Long
        
        Call .Add(gInt.Of(-2))
        Call .Add(gInt.Of(-1))
     
        For i = 1 To 15
            Call .Add(gInt.Of(i))
        Next
        Call .Add(gInt.Of(10))
        Call .Add(gInt.Of(10))
        Call .Add(gInt.Of(15))
        Call .Add(gInt.Of(27))
        Call .Add(gInt.Of(28))
    
        Call .Shuffle
        Call .Sort(, IGenericValue)
        Debug.Print .BinarySearch(gInt.Of(17))
    End With
    
    i = 0
     
    With List.Range
        Do While .HasNext
            If .Current Is Nothing Then
                Debug.Print i & " :: Nothing"
            Else
                Debug.Print i & " :: " & .Current.Instance.ToString
            End If
            i = i + 1
        Loop
    
    End With
    
    Dim SortedList As GenericSortedList
    If SortedList Is Nothing Then
        Set SortedList = GenericSortedList.BuildFrom(Sequence:=List, AscendingOrdered:=IGenericValue)
        Call SortedList.Add(gInt.Of(0))
    End If
    
    Debug.Print SortedList.IndexOf(gInt.Of(17))
    
    i = 0
    
    With SortedList.Elements.Iterator
        Do While .HasNext
            If .Current Is Nothing Then
                Debug.Print i & " :: Nothing"
            Else
                Debug.Print i & " :: " & .Current.Instance.ToString
            End If
            i = i + 1
        Loop
    End With
    
End Sub

Sub TestSQl()
    With GenericSqlManager
        Debug.Print .UpdateTable("Test", GenericSequence.Of(gString.Of("F1"), gString.Of("F2"), gString.Of("F3")).ToArray).ToString
    End With
End Sub

