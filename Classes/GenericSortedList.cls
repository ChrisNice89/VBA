VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSortedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "Collections"
Option Explicit

'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericSortedList
'*
'* Purpose: The GenericSortedList class represents a collection of elements that are sorted
'*          (based on the associated Comparer implementation) with O(log n) retrieval, where n is the number of elements in the List and are accessible by index.
'*          For example, if the values are of primitive types, then the SortedList is sorted in ascending order of the elements.
'*          The index sequence is based on the sort sequence. When an element is added, it is inserted into SortedList in the correct sort order, and the indexing adjusts accordingly.
'*          When an element is removed, the indexing also adjusts accordingly. Therefore, the index of a specific element might change as elements are added or removed from the SortedList object.
'*          The capacity of a SortedList object is the number of elements the SortedList can hold.
'*          As elements are added to a SortedList, the capacity is automatically increased as required through reallocation.
'*          The capacity(is never automatically decreased) can be decreased by calling TrimExcess or by setting the Capacity property explicitly.
'*          Operations on a SortedList object tend to be slower than operations on a GenericOrderedList object because of the sorting.
'*          However, the SortedList offers better performance on search operations.
'*
'* Notes:   A SortedList object internally maintains an array to store the elements of the list
'*          It maintains ascending order (according to the associated Comparer implementation) and does allow duplicate elements.
'*          An element can be accessed by its index.
'*          Uses binary search to retrieve items
'*          Supports efficient indexed retrieval of elements
'*          Uses less memory than SortedSet.
'*          If the list is populated all at once from sorted data, SortedList is faster than SortedSet.
'*
'*          If NULL values are expected in the collection, a NULL comparison operator must be provided. Otherwise null pointer exceptions can occur!
'*
'*
'*======================================================================================================================

Implements IGeneric
Implements IGenericCollection
Implements IGenericSequence

#If Win64 Then
    Private Const POINTERSIZE As LongPtr = 8
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dst As Any, ByRef Src As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef Dst As Any, ByVal Length As LongPtr, ByVal Fill As Byte)
    Private Declare PtrSafe Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As LongPtr, ByRef Dst As LongPtr, ByRef Src As LongPtr) As Long
    Private Declare PtrSafe Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Length As Long, ByRef Dst As LongPtr, ByVal Src As LongPtr) As Long
    Private Declare PtrSafe Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As LongPtr
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
    Private Const POINTERSIZE As Long = 4
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dst As Any, ByRef Src As Any, ByVal Length As Long)
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal Length As Long)
    Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef Dst As Any, ByVal Length As Long, ByVal Fill As Byte)
    Private Declare Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As Long, ByRef Dst As Long, ByRef Src As Long) As Long
    Private Declare Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Length As Long, ByRef Dst As Long, ByVal Src As Long) As Long
    Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

Private Const DEFAULT_CAPACITY As Long = 8
Private Const NOT_SPECIFIED As Long = -1 'Dont change!
Private Const SORT_ORDER As Long = SortOrder.Ascending
Private Const COPYMEMORY_THRESHOLD As Long = 32 'Depends on machine

Private Type Member
    Items As GenericArray
    AscendingOrdered As IGenericPredicator
    Count As Long
End Type
Private This As Member

Public Property Get Stream() As GenericSequence: Set Stream = GenericSequence.Stream(Me): End Property
Public Property Get Instance() As IGeneric: Set Instance = Me: End Property
Public Property Get Elements() As IGenericCollection: Set Elements = Me: End Property
Public Property Get First() As Long: First = This.Items.LowerBound: End Property
Public Property Get Last() As Long
    
    Last = This.Items.LowerBound
    
    If This.Count > 0 Then _
        Last = This.Count - Last - 1
   
End Property

Public Property Get Direction() As SortOrder: Direction = SORT_ORDER: End Property
Public Property Get AscendingOrdered() As IGenericPredicator: Set AscendingOrdered = This.AscendingOrdered: End Property

Public Property Get Capacity() As Long: Capacity = This.Items.Length: End Property
Public Property Let Capacity(ByVal Value As Long)
    
    Dim tmp As GenericArray
    
    If (Value < 0) Then _
        Call GenericError.Build(Me, "Public Property Let Capacity").ArgOutOfRange("(Value < 0)")
        
    If (This.Items Is Nothing) Then _
        Call GenericError.Build(Me, "Public Property Let Capacity").IsNothing("This.Items Is Nothing")
    
    If (Value > This.Items.Length) Then
        
        Set tmp = GenericArray.Build(Value)
        
        If (This.Count > 0) Then _
            Call This.Items.CopyTo(Target:=tmp, Index:=tmp.LowerBound, BeginAt:=This.Items.LowerBound, Count:=This.Count)

        Set This.Items = tmp

    End If
    
End Property

Public Property Get GetAt(ByVal Index As Long) As Object
    
    If (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Public Property Get GetAt").IndexOutofRange("(Index >= This.Count)")
    
    Set GetAt = This.Items.ElementAt(Index)

End Property

Public Function Build(Optional ByVal Capacity As Long, Optional ByVal AscendingOrdered As IGenericPredicator) As GenericSortedList
    
    Dim newList As Member
    
    If (Capacity < 0) Then _
        Call GenericError.Build(Me, "Build").ArgOutOfRange("Capacity < 0")
    
    Set newList.AscendingOrdered = AscendingOrdered
    
    'Avoid null
    Set newList.Items = GenericArray
    If (Capacity > 0) Then _
        Set newList.Items = GenericArray.Build(Capacity)

    Set Build = New GenericSortedList
    Call System.Inject(Build, VarPtr(newList), LenB(newList))
    
End Function

Public Function BuildFrom(ByVal Sequence As IGenericSequence, Optional ByVal AscendingOrdered As IGenericPredicator) As GenericSortedList: Set BuildFrom = GenericSortedList.Build(AscendingOrdered:=AscendingOrdered).AddAll(Sequence): End Function

Public Function EnsureCapacity(ByVal Min As Long) As GenericSortedList
    
    Dim Size As Long
    
    With This.Items
        If (.Length < Min) Then
            Size = VBA.IIf(.Length = 0, Min, .Length * 2)
            
            If (Size > .MaxLength) Then _
                Size = .MaxLength
            
            If (Size < Min) Then _
                Size = Min
        
            Me.Capacity = Size
            
        End If
    End With
    
    Set EnsureCapacity = Me
    
End Function

Private Sub Expand()

    Dim Size As Long
    
    With This.Items
        
        Size = VBA.IIf(.Length = 0, DEFAULT_CAPACITY, .Length * 2)
        
        If (Size > .MaxLength) Then _
            Size = .MaxLength
    
        Me.Capacity = Size
            
    End With
    
End Sub

'SortedList.Clear();
'SortedList.TrimExcess();
Public Function TrimExcess() As GenericSortedSet
     
    If This.Count < VBA.Int(CDbl(This.Items.Length) * 0.9) Then _
        Me.Capacity = This.Count
    
    Set TrimExcess = Me
    
End Function

Public Function Contains(ByVal Element As Object) As Boolean: Contains = Not (Me.IndexOf(Element) = -1): End Function

Public Function ContainsAll(ByVal Sequence As IGenericSequence) As Boolean
    
    '@Ignore VariableNotAssigned
    Dim Candidate As Object
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function ContainsAll").IsNothing("Sequence")
    
    If (This.Count = 0) Then _
        Exit Function
        
    If Not (Sequence Is Me) Then
        With Sequence.Iterator
            Do While .HasNext(Candidate)
                If (Me.IndexOf(Candidate) = -1) Then _
                    Exit Function
            Loop
        End With
    End If
    
    ContainsAll = True

End Function

'Returns the index of the entry in this sorted list. The
'entry is located through a binary search, and thus the average execution
'time of this method is proportional to Log2(Size), where
'Size is the Size of this sorted list. The returned entry is -1 if
'the given entry does not occur in this sorted list. Null is an valid
'entry value.
Public Function IndexOf(ByVal Element As Object) As Long

    Dim i As Long
    Dim Casted As IGenericValue
    
    i = This.Items.BinarySearch(Element:=Element, Direction:=SORT_ORDER, AscendingOrdered:=AscendingOrdered, Length:=This.Count)
    
    If (i = This.Count) Then
        IndexOf = -1
        Exit Function
    End If
    
    If (SORT_ORDER = SortOrder.Ascending) Then
    
        If (AscendingOrdered Is Nothing) Then
            Set Casted = Element
            If Casted.IsGreaterThan(This.Items.ElementAt(i)) Then _
                i = -1
                
        ElseIf AscendingOrdered.IsTrue(Element, This.Items.ElementAt(i)) Then
            i = -1
        
        End If
        
    Else
    
        If (AscendingOrdered Is Nothing) Then
            Set Casted = Element
            If Casted.IsLessThan(This.Items.ElementAt(i)) Then _
                i = -1
                
        ElseIf AscendingOrdered.IsTrue(This.Items.ElementAt(i), Element) Then
            i = -1
        
        End If
    
    End If
    
    IndexOf = i
    
End Function

Public Function Distinct(Optional ByVal Predicate As IGenericPredicator) As GenericOrderedList

    This.Count = This.Items.Distinct(Predicate:=Predicate, Index:=This.Items.LowerBound, Length:=This.Count)
    Set Distinct = Me
    
End Function

Public Function TakeAt(ByVal Index As Long) As Object

    If (Index < This.Items.LowerBound) Or (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Public Function TakeAt").IndexOutofRange("(Index < This.Items.LowerBound) Or (Index >= This.Count)")
    
    Set TakeAt = This.Items.ElementAt(Index)
    Call Me.RemoveAt(Index)
    
End Function

Public Function TryRemove(ByVal Element As Object) As Boolean

    Dim i As Long
    
    i = Me.IndexOf(Element)
       
    If (i = -1) Then _
        Exit Function
        
    Call Me.RemoveAt(i)
    TryRemove = True
    
End Function

Public Sub RemoveAt(ByVal Index As Long)
    
    Dim i As Long, Length As Long
    
    If (Index < This.Items.LowerBound Or Index >= This.Count) Then _
        Call GenericError.Build(Me, "Public Sub RemoveAt").ArgOutOfRange("(Index >= This.Count)")
   
    With This.Items

        Set .ElementAt(Index) = Nothing
        
        If (Index < InterlockedDecrement(This.Count)) Then
            Length = This.Count - Index
            If (Length <= COPYMEMORY_THRESHOLD) Then
            
                For i = Index To This.Count - 1: Set .ElementAt(i) = .ElementAt(i + 1): Next
            
            Else
            
                Call CopyMemory(Dst:=ByVal .AdressOf(Index), Src:=ByVal .AdressOf(Index + 1), Length:=.SizeOf(Length)): _
                Call ZeroMemory(Dst:=ByVal .AdressOf(This.Count), Length:=.SizeOf(1))
            
            End If
        End If
        
    End With
    
End Sub

Public Function RetainAll(ByVal Collection As IGenericCollection) As Long: RetainAll = BatchRemove(Collection:=Collection, Complement:=True): End Function
Public Function RemoveAll(ByVal Collection As IGenericCollection) As Long: RemoveAll = BatchRemove(Collection:=Collection, Complement:=False): End Function

Public Function Slice(ByVal BeginAt As Long, Optional ByVal Count As Long = NOT_SPECIFIED) As GenericSortedList
    
    Dim newList As Member
    
    If (BeginAt < This.Items.LowerBound) Then _
        Call GenericError.Build(Me, "Public Function Slice").ArgOutOfRange("(BeginAt < This.Items.LowerBound)")
    
    If (Count = NOT_SPECIFIED) Then _
        Count = This.Count - BeginAt
   
    If (Count <= 0) Then _
        Call GenericError.Build(Me, "Public Function Slice").ArgOutOfRange("(Count <= 0)")
   
    If (This.Count - BeginAt < Count) Then _
        Call GenericError.Build(Me, "Public Function Slice").InvalidInput("(This.Count - BeginAt < Count)")
   
    Set newList.Items = GenericArray.Build(Count)
    Call This.Items.CopyTo(Target:=newList.Items, Index:=newList.Items.LowerBound, BeginAt:=BeginAt, Count:=Count)

    newList.Count = Count
    Set newList.AscendingOrdered = AscendingOrdered
    
    Set Slice = New GenericSortedList
    Call System.Inject(Slice, VarPtr(newList), LenB(newList))
    
End Function

Public Function Range(Optional ByVal BeginAt As Long = NOT_SPECIFIED, Optional ByVal Count As Long = NOT_SPECIFIED) As IGenericIterator
        
    If (This.Items Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Range").IsNothing("This.Items")
        
    If (BeginAt = NOT_SPECIFIED) Then
        BeginAt = This.Items.LowerBound
    ElseIf (BeginAt < 0) Then
        Call GenericError.Build(Me, "Public Function Range").ArgOutOfRange("BeginAt < 0")
    End If
    
    If (Count = NOT_SPECIFIED) Then
        Count = This.Count - BeginAt
    ElseIf (Count < 0) Then
        Call GenericError.Build(Me, "Public Function Range").ArgOutOfRange("Count < 0")
    End If
    
    If (This.Count - BeginAt < Count) Then _
        Call GenericError.Build(Me, "Public Function Range").InvalidInput("(This.Count - BeginAt < Count)")
        
    Set Range = This.Items.Range(BeginAt:=BeginAt, Count:=Count)
   
End Function

Public Function Add(ByVal Element As Object) As Long
       
    Dim i As Long
   
    i = This.Items.BinarySearch(Element:=Element, Direction:=SORT_ORDER, AscendingOrdered:=AscendingOrdered, Length:=This.Count)
    Call Insert(i, Element)
    
    Add = i
    
End Function

Public Function AddAll(ByVal Sequence As IGenericSequence) As GenericSortedList

    '@Ignore VariableNotAssigned
    Dim Element As Object
    Dim Collection As IGenericCollection, tmp As GenericArray
    Dim Count As Long
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Sub AddAll").IsNothing("Sequence")
    
    Set AddAll = Me
    
    If TypeOf Sequence Is IGenericCollection Then
        Set Collection = Sequence
        Count = Collection.Count
        
        If (Count = 0) Then _
            Exit Function
     
        If TypeOf Collection Is GenericSortedList Then
            If IsComparable(Me, Collection) Then
        
                Set tmp = GenericArray.Build(Count)
                Call Collection.CopyTo(tmp, tmp.LowerBound)
                
                If (This.Count > 0) Then
                    Set This.Items = Merge(This.Items, This.Count, tmp, Count, AscendingOrdered, SORT_ORDER)
                Else
                    Set This.Items = tmp
                End If
                
                This.Count = This.Items.Length
                Exit Function
            End If
        End If
    
        Call Me.EnsureCapacity(This.Count + Count)
                
   End If
   
    'The hard way
    With Sequence.Iterator
        Do While .HasNext(Element): Call Insert(This.Items.BinarySearch(Element:=Element, Direction:=SORT_ORDER, AscendingOrdered:=AscendingOrdered, Length:=This.Count), Element): Loop
    End With
    
End Function

'TODO Check comparer Equality
Private Function IsComparable(ByVal A As GenericSortedList, ByVal B As GenericSortedList) As Boolean
    IsComparable = (A.Direction = B.Direction And A.AscendingOrdered Is B.AscendingOrdered)
End Function

Private Sub Insert(ByVal Index As Long, ByVal Element As Object)
    
    Dim i As Long, Length As Long
    
'    If (Index > This.Count) Then _
'        Call GenericError.Build(Me, "Public Sub Insert").IndexOutofRange("(Index > This.Count)")
        
    If (This.Count = This.Items.Length) Then _
        Call Expand
    
    Length = This.Count - Index
    
    With This.Items
    
        If (Length > 0) Then
            If (Length <= COPYMEMORY_THRESHOLD) Then
                
                For i = This.Count To Index + 1 Step -1: Set .ElementAt(i) = .ElementAt(i - 1): Next
            
            Else
            
                Call CopyMemory(Dst:=ByVal .AdressOf(Index + 1), Src:=ByVal .AdressOf(Index), Length:=.SizeOf(Length)): _
                Call ZeroMemory(Dst:=ByVal .AdressOf(Index), Length:=.SizeOf(1))
            
            End If
        End If
        
        Set .ElementAt(Index) = Element
        
    End With
    
    This.Count = This.Count + 1

End Sub

'Pops an item from the top of the stack.  If the stack is empty, Pop
'throws an InvalidOperationException.
Public Function Pop() As Object
    
    Set Pop = This.Items.ElementAt(This.Count - 1)
    Call Me.RemoveAt(This.Count - 1)
    
End Function

'Pushes an item to the top of the stack.
Public Function Push(ByVal Element As Object) As GenericSortedList
    Call Me.Add(Element:=Element)
    Set Push = Me
End Function
'Returns the top object on the stack without removing it.  If the stack
'is empty, Peek throws an InvalidOperationException.
Public Function Peek() As Object: Set Peek = Me.GetAt(Me.Last): End Function
                                                                                                 
Private Function BatchRemove(ByVal Collection As IGenericCollection, ByVal Complement As Boolean) As Long
    
    Dim Candidate As Object
    Dim Result As Long
    Dim r As Long
    Dim w As Long
    
    If (Collection Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function RemoveAll").IsNothing("(Collection Is Nothing)")
    
    If (Collection.Count > 0) Then
        With This.Items
            Do While (r < This.Count)
                Set Candidate = .ElementAt(r)
                If (Collection.Contains(Candidate) = Complement) Then
                    Set .ElementAt(w) = Candidate
                    w = w + 1
                End If
                r = r + 1
            Loop
            
            If Not (r = This.Count) Then _
                Call GenericError.Build(Me, "Public Sub RemoveAll").InvalidOperation
            
            For r = w To This.Count - 1: Set .ElementAt(r) = Nothing: Next
            
            Result = This.Count - w
           
        End With
    End If
    
    This.Count = This.Count - Result
    BatchRemove = Result
    
End Function
'todo Comparer
Private Function Merge( _
                        ByVal SortedA As GenericArray, _
                        ByVal LengthA As Long, _
                        ByVal SortedB As GenericArray, _
                        ByVal LengthB As Long, _
                        Optional ByVal Comparer As IGenericPredicator, _
                        Optional ByVal Direction As SortOrder = SORT_ORDER) As GenericArray
    
    Dim i As Long, j As Long, k As Long
    Dim A As Object, B As Object
 
    Set Merge = GenericArray.Build(LengthA + LengthB)
    
    If (Comparer Is Nothing) Then _
        Set Comparer = IGenericValue
 
    With Merge
        
        If (Direction = Ascending) Then
        
            Do While (i < LengthA And j < LengthB)
            
                Set A = SortedA.ElementAt(i)
                Set B = SortedB.ElementAt(j)
    
                If AscendingOrdered.IsTrue(A, B) Then
                    Set .ElementAt(k) = A: i = i + 1
                Else
                    Set .ElementAt(k) = B: j = j + 1
                End If
    
                k = k + 1
            Loop
            
        Else
            
            Do While (i < LengthA And j < LengthB)
        
                Set A = SortedA.ElementAt(i)
                Set B = SortedB.ElementAt(j)
    
                If AscendingOrdered.IsTrue(B, A) Then
                    Set .ElementAt(k) = A: i = i + 1
                Else
                    Set .ElementAt(k) = B: j = j + 1
                End If
    
                k = k + 1
            Loop
        
        End If
        
        Do While (i < LengthA)
            Set .ElementAt(k) = SortedA.ElementAt(i)
            i = i + 1
            k = k + 1
        Loop

        Do While (j < LengthB)
            Set .ElementAt(k) = SortedB.ElementAt(j)
            j = j + 1
            k = k + 1
        Loop
            
    End With

End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericSortedList: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = GenericSortedList: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(This): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(This): End Property
Private Property Get IGeneric_HashCode() As Long:  IGeneric_HashCode = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericSortedList):  End Property

'@Ignore NonReturningFunction, EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As Object) As Boolean: IGeneric_Equals = Me Is Other: End Function

Private Function IGeneric_ToString() As String
    With This
        IGeneric_ToString = _
                            "Class: " & VBA.TypeName$(Me) & vbNewLine & _
                            "Capacity: " & .Items.Length & vbNewLine & _
                            "Elements: " & .Count & vbNewLine & _
                            "Sortorder: " & VBA.IIf(SORT_ORDER = 1, "Ascending", "Descending")
    End With
End Function

'IGenericCollection
Private Property Get IGenericCollection_HasRandomAccess() As Boolean: IGenericCollection_HasRandomAccess = True: End Property
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = This.Count: End Property

Private Function IGenericCollection_GetAt(ByVal Index As Long) As Object

    If (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Private Function IGenericCollection_GetAt").IndexOutofRange("(Index >= This.Count)")
    
    Set IGenericCollection_GetAt = This.Items.ElementAt(Index)

End Function

Private Function IGenericCollection_Contains(ByVal Element As Object) As Boolean: IGenericCollection_Contains = Not (Me.IndexOf(Element) = -1): End Function

'ToArray returns a new Object array containing the contents of the List.
'This requires copying the List, which is an O(n) operation.
Private Function IGenericCollection_ToArray() As GenericArray

    Set IGenericCollection_ToArray = GenericArray 'avoid null
    
    If (This.Count = 0) Then _
        Exit Function
        
    Set IGenericCollection_ToArray = GenericArray.Build(This.Count)
    Call Me.Elements.CopyTo(Target:=IGenericCollection_ToArray, Index:=IGenericCollection_ToArray.LowerBound)
    
End Function

Private Sub IGenericCollection_CopyTo(ByVal Target As GenericArray, ByVal Index As Long)
    
    If (This.Count = 0) Then _
        Exit Sub
        
    Call This.Items.CopyTo(Target:=Target, Index:=Index, BeginAt:=This.Items.LowerBound, Count:=This.Count)
    
End Sub

Private Sub IGenericCollection_Clear()
        
    If (This.Count = 0) Then _
        Exit Sub
 
    Call This.Items.Elements.Clear
    This.Count = 0

End Sub

Private Function IGenericCollection_Copy() As IGenericCollection
    
    Dim Clone As GenericSortedList
    
    Set Clone = GenericSortedList.Build(AscendingOrdered:=AscendingOrdered)
    Call Clone.AddAll(Me)
   
    Set IGenericCollection_Copy = Clone
    
End Function

Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Range: End Function

'IGenericSequence_Iterator
Private Function IGenericSequence_Iterator() As IGenericIterator: Set IGenericSequence_Iterator = Me.Range: End Function


