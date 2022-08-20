VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericOrderedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@PredeclaredId
'@IgnoreModule FunctionReturnValueAlwaysDiscarded, VariableNotAssigned, NonReturningFunction, UnassignedVariableUsage, ProcedureNotUsed, SuspiciousPredeclaredInstanceAccess, ConstantNotUsed, AssignedByValParameter
'@Folder "Collections"
Option Explicit

'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericOrderedList
'*
'* Purpose: The GenericOrderedList is a collection which maintains an ordering for its elements.
'*          that can be accessed by index and having methods for sorting, searching, and modifying list.
'*          i.e. GenericOrderedList grows dynamically as the elements are added to it.
'*          If the Size of the current elements (including the new element to be added to the GenericOrderedList)
'*          is greater than the maximum Size of the array then increase the Size of array.
'*          But the Size of the array can not be increased dynamically.
'*          So, what happens internally is, a new Array is created and
'*          the old array is copied into the new array.
'*          The count,  get, set, and iterator operations run in constant time.
'*          The add operation runs in amortized constant time, that is, adding n elements requires O(n) time.
'*          All of the other operations run in linear time (roughly speaking).
'*          The constant factor is low compared to that for a LinkedList implementation.
'*          This class provides methods to manipulate the size of the array that is used internally to store the list.
'*          (List class is roughly equivalent to a Vector)
'*          Each List instance has a capacity. The capacity is the size of the array used to store the elements in the list.
'*          It is always at least as large as the list size. As elements are added to an List, its capacity grows automatically.
'*          The details of the growth policy are not specified beyond the fact that adding an element has constant amortized time cost.
'*          An application can increase the capacity of a List instance before adding a large number of elements using the
'*          Let Capacity Property. This may reduce the amount of incremental reallocation.
'*
'* Notes:   The important points about GenericOrderedList class are:
'*          List is based on GenericArray data structure.
'*          List class can contain duplicate elements.
'*          List class maintains insertion order.
'*          List allows random access because array works at the index basis.
'*          List, manipulation is little bit slower than a LinkedList because a lot of shifting needs
'*          to occur if any element is removed from the list.
'*          Unlike sets, lists typically allow duplicate elements.
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
    Private Declare PtrSafe Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As LongPtr, ByRef Dst As LongPtr, ByVal Src As LongPtr) As Long
    Private Declare PtrSafe Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Length As Long, ByRef Dst As LongPtr, ByVal Src As LongPtr) As Long
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
    Private Const POINTERSIZE As Long = 4
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dst As Any, ByRef Src As Any, ByVal Length As Long)
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal Length As Long)
    Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef Dst As Any, ByVal Length As Long, ByVal Fill As Byte)
    Private Declare Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As Long, ByRef Dst As Long, ByVal Src As Long) As Long
    Private Declare Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Length As Long, ByRef Dst As Long, ByVal Src As Long) As Long
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

Private Const DEFAULT_CAPACITY As Long = 8
Private Const NOT_SPECIFIED As Long = -1 'Dont change!
Private Const SORT_ORDER As Long = SortOrder.Ascending
Private Const COPYMEMORY_THRESHOLD As Long = 32 'Depends on machine

Private Type Member
    Items As GenericArray
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

Public Property Set ElementAt(ByVal Index As Long, ByVal Element As Object)
   
    If (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Public Property Set ElementAt").IndexOutofRange("(Index >= This.Count)")
    
    Set This.Items.ElementAt(Index) = Element

End Property

Public Property Get ElementAt(ByVal Index As Long) As Object
    
    If (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Public Property Get ElementAt").IndexOutofRange("(Index >= This.Count)")
    
    Set ElementAt = This.Items.ElementAt(Index)

End Property

Public Property Get GetAt(ByVal Index As Long) As Object
    
    If (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Public Property Get GetAt").IndexOutofRange("(Index >= This.Count)")
    
    Set GetAt = This.Items.ElementAt(Index)

End Property

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

Public Function Build(Optional ByVal Capacity As Long) As GenericOrderedList
    
    If (Capacity < 0) Then _
        Call GenericError.Build(Me, "Build").ArgOutOfRange("Capacity < 0")
    
    Set Build = New GenericOrderedList
    If (Capacity > 0) Then _
        Call Build.EnsureCapacity(Capacity)

End Function

Private Sub Class_Initialize(): Set This.Items = GenericArray: End Sub

Public Function BuildFrom(ByVal Sequence As IGenericSequence) As GenericOrderedList: Set BuildFrom = GenericOrderedList.Build().AddAll(Sequence): End Function

'@Ignore FunctionReturnValueAlwaysDiscarded
Public Function EnsureCapacity(ByVal Min As Long) As GenericOrderedList
    
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

Public Sub Swap(ByVal i As Long, ByVal j As Long)

    Dim Element As Object
    
    If (i >= This.Count) Or (j >= This.Count) Then _
        Call GenericError.Build(Me, "Public Sub Swap").IndexOutofRange("(i >= This.Count) Or (j >= This.Count)")
     
    With This.Items
        Set Element = .ElementAt(i): Set .ElementAt(i) = .ElementAt(j): Set .ElementAt(j) = Element
    End With
    
End Sub

'Pops an item from the top of the stack.  If the stack is empty, Pop
'throws an InvalidOperationException.
Public Function Pop() As Object
    Set Pop = This.Items.ElementAt(This.Count - 1)
    Call Me.RemoveAt(This.Count - 1)
End Function

'Pushes an item to the top of the stack.
Public Function Push(ByVal Element As Object) As GenericOrderedList
    Call Me.Add(Element)
    Set Push = Me
End Function
'Returns the top object on the stack without removing it.  If the stack
'is empty, Peek throws an InvalidOperationException.
Public Function Peek() As Object: Set Peek = Me.ElementAt(Me.Last): End Function

Public Sub Add(ByVal Element As Object)
    
    If (This.Count = This.Items.Length) Then _
        Call Expand
    
    Set This.Items.ElementAt(This.Count) = Element
    This.Count = This.Count + 1
 
End Sub

Public Function AddAll(ByVal Sequence As IGenericSequence) As GenericOrderedList: Set AddAll = Me.InsertAll(Index:=This.Count, Sequence:=Sequence): End Function

Public Function InsertAll(ByVal Index As Long, ByVal Sequence As IGenericSequence) As GenericOrderedList

    Dim Element As Object
    Dim Collection As IGenericCollection
    Dim Count As Long, i As Long, Length As Long
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Sub InsertAll").IsNothing("Sequence")

    If (Index > This.Count) Then _
        Call GenericError.Build(Me, "Public Sub InsertAll").IndexOutofRange("(Index > This.Count)")
    
    Set InsertAll = Me
    
    If TypeOf Sequence Is IGenericCollection Then
        '@Ignore SetAssignmentWithIncompatibleObjectType
        Set Collection = Sequence
        Count = Collection.Count
        
        If (Count = 0) Then _
            Exit Function
            
        Call Me.EnsureCapacity(This.Count + Count)
        
        If (Collection Is Me) Then _
            Set Collection = Collection.Copy
            
        Length = This.Count - Index
        
        If (Length > 0) Then
            With This.Items

                If (Length <= COPYMEMORY_THRESHOLD) Then

                    For i = This.Count To Index + 1 Step -1: Set .ElementAt(i + Count - 1) = .ElementAt(i - 1): Next

                Else

                    Call CopyMemory(Dst:=ByVal .AdressOf(Index + Count), Src:=ByVal .AdressOf(Index), Length:=.SizeOf(Length)): _
                    Call ZeroMemory(Dst:=ByVal .AdressOf(Index), Length:=.SizeOf(Count))

                End If

            End With
        End If
        
        Call Collection.CopyTo(Target:=This.Items, Index:=Index)
        This.Count = This.Count + Count
        Exit Function
       
    End If
  
    With Sequence.Iterator
        If (Index < This.Count) Then
            Do While .HasNext(Element): Call Me.Insert(Index, Element): Index = Index + 1: Loop
        Else
            Do While .HasNext(Element): Call Me.Add(Element): Loop
        End If
    End With

End Function

Public Sub Insert(ByVal Index As Long, ByVal Element As Object)
        
    Dim i As Long, Length As Long
    
    If (Index > This.Count) Then _
        Call GenericError.Build(Me, "Public Sub Insert").IndexOutofRange("(Index > This.Count)")

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

Public Function SelectKth(ByVal Kth As Long, Optional ByVal Direction As SortOrder = SORT_ORDER, Optional ByVal AscendingOrdered As IGenericPredicator) As Object
    
    Kth = This.Items.SelectKth(Kth:=Kth, Direction:=Direction, AscendingOrdered:=AscendingOrdered, Length:=This.Count)
    Set SelectKth = This.Items.ElementAt(Kth)
    
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

Public Function TryRemove(ByVal Element As Object, Optional ByVal Predicate As IGenericPredicator) As Boolean

    Dim i As Long
    i = This.Items.IndexOf(Element:=Element, Length:=This.Count, Predicate:=Predicate)
    
    If (i < 0) Then _
        Exit Function
        
    Call Me.RemoveAt(i)
    TryRemove = True
    
End Function

Public Sub RemoveAt(ByVal Index As Long)
    
    Dim i As Long, Length As Long
    
    If (Index < This.Items.LowerBound) Or (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Public Sub RemoveAt").IndexOutofRange("(Index < This.Items.LowerBound) Or (Index >= This.Count)")
    
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

Public Function Slice(ByVal BeginAt As Long, Optional ByVal Count As Long = NOT_SPECIFIED) As GenericOrderedList
    
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

    Set Slice = New GenericOrderedList
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

Public Function Fill(ByVal Instance As Object, Optional ByVal Index As Long = NOT_SPECIFIED, Optional ByVal Count As Long = NOT_SPECIFIED) As GenericOrderedList
    
    Dim i As Long
    
    If (Index = NOT_SPECIFIED) Then
        Index = This.Items.LowerBound
    ElseIf (Index < 0) Then
        Call GenericError.Build(Me, "Public Function Fill").ArgOutOfRange("Index < 0")
    End If
    
    If (Count = NOT_SPECIFIED) Then
        Count = This.Count
    ElseIf (Count < 0) Then
        Call GenericError.Build(Me, "Public Function Fill").ArgOutOfRange("Count < 0")
    End If
    
    If (This.Count - Index < Count) Then _
        Call GenericError.Build(Me, "Public Sub Fill").InvalidInput("(This.Count - Index < Count)")
   
    With This.Items
        For i = Index To Index + Count: Set This.Items.ElementAt(i) = Instance: Next
    End With
    
    Set Fill = Me
    
End Function

Public Function Sort( _
                       Optional ByVal Direction As SortOrder = SORT_ORDER, _
                       Optional ByVal AscendingOrdered As IGenericPredicator, _
                       Optional ByVal Index As Long = NOT_SPECIFIED, _
                       Optional ByVal Count As Long = NOT_SPECIFIED) As GenericOrderedList

    If (Index = NOT_SPECIFIED) Then
        Index = This.Items.LowerBound
    ElseIf (Index < 0) Then
        Call GenericError.Build(Me, "Public Function Sort").ArgOutOfRange("Index < 0")
    End If
    
    If (Count = NOT_SPECIFIED) Then
        Count = This.Count
    ElseIf (Count < 0) Then
        Call GenericError.Build(Me, "Public Function Sort").ArgOutOfRange("Count < 0")
    End If
    
    If (This.Count - Index < Count) Then _
        Call GenericError.Build(Me, "Public Sub Sort").InvalidInput("(This.Count - Index < Count)")
        
    Call This.Items.Sort(Direction:=Direction, AscendingOrdered:=AscendingOrdered, Index:=Index, Length:=Count)
    Set Sort = Me
    
End Function

Public Function Contains(ByVal Element As Object, Optional ByVal Predicate As IGenericPredicator) As Boolean: Contains = Not (This.Items.IndexOf(Element:=Element, Length:=This.Count, Predicate:=Predicate) = -1): End Function

'Searches a section of the list for a given element using a binary search
'algorithm. Elements of the list are compared to the search value using the IGeneric
'comparer, which in that case must be implemented by all elements of the
'list and the given search value. This method assumes that the given
'section of the list is already sorted; if this is not the case, the
'result will be incorrect.
'
'The method returns the index of the given value in the list. If the
'list does not contain the given value, the method returns a negative
'integer. The bitwise complement operator (~) can be applied to a
'negative result to produce the index of the first element (if any) that
'is larger than the given search value. This is also the index at which
'the search value should be inserted into the list in order for the list
'to remain sorted.
Public Function BinarySearch( _
                                ByVal Element As Object, _
                                Optional ByVal Direction As SortOrder = SORT_ORDER, _
                                Optional ByVal AscendingOrdered As IGenericPredicator) As Long

    
    Dim i As Long
    Dim Casted As IGenericValue
    
    i = This.Items.BinarySearch(Element:=Element, Direction:=Direction, AscendingOrdered:=AscendingOrdered, Length:=This.Count)
    
    If (i = This.Count) Then
        BinarySearch = Not i
        Exit Function
    End If
    
    If (SORT_ORDER = SortOrder.Ascending) Then
    
        If (AscendingOrdered Is Nothing) Then
            Set Casted = Element
            If Casted.IsGreaterThan(This.Items.ElementAt(i)) Then _
                i = Not i
                
        ElseIf AscendingOrdered.IsTrue(Element, This.Items.ElementAt(i)) Then
            i = Not i
        
        End If
        
    Else
    
        If (AscendingOrdered Is Nothing) Then
            Set Casted = Element
            If Casted.IsLessThan(This.Items.ElementAt(i)) Then _
                i = Not i
                
        ElseIf AscendingOrdered.IsTrue(This.Items.ElementAt(i), Element) Then
            i = Not i
        
        End If
    
    End If
    
    BinarySearch = i
    
End Function

Public Function IndexOf(ByVal Element As Object, Optional ByVal Predicate As IGenericPredicator) As Long: IndexOf = This.Items.IndexOf(Element:=Element, Length:=This.Count, Predicate:=Predicate): End Function

Public Function Reverse( _
                            Optional ByVal Index As Long = NOT_SPECIFIED, _
                            Optional ByVal Count As Long = NOT_SPECIFIED) As GenericOrderedList
    
    If (Index = NOT_SPECIFIED) Then
        Index = This.Items.LowerBound
    ElseIf (Index < 0) Then
        Call GenericError.Build(Me, "Public Function Reverse").ArgOutOfRange("Index < 0")
    End If
    
    If (Count = NOT_SPECIFIED) Then
        Count = This.Count
    ElseIf (Count < 0) Then
        Call GenericError.Build(Me, "Public Function Reverse").ArgOutOfRange("Count < 0")
    End If
    
    If (This.Count - Index < Count) Then _
        Call GenericError.Build(Me, "Public Function Reverse").InvalidInput("(This.Count - Index < Count)")
        
    Call This.Items.Reverse(Index:=Index, Length:=Count)
    Set Reverse = Me
    
End Function

Public Function Shuffle() As GenericOrderedList
    
    Call This.Items.Shuffle(This.Items.LowerBound, This.Count)
    Set Shuffle = Me
    
End Function

'Sets the capacity of this list to the Size of the list. This method can
'be used to minimize a list's memory overhead once it is known that no
'new elements will be added to the list. To completely clear a list and
'release all memory referenced by the list, execute the following
'statements:
'
'list.Clear();
'list.TrimExcess();
Public Function TrimExcess() As GenericOrderedList
     
    If This.Count < VBA.Int(CDbl(This.Items.Length) * 0.9) Then _
        Me.Capacity = This.Count
    
    Set TrimExcess = Me
    
End Function

Private Function BatchRemove(ByVal Collection As IGenericCollection, ByVal Complement As Boolean) As Long
    
    Dim Candidate As Object
    
    Dim Result As Long
    Dim r As Long
    Dim w As Long
    
    If (Collection Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function RemoveBulk").IsNothing("(Collection Is Nothing)")
    
    If (Collection.Count > 0) Then
'        Set Sequence = Collection.Iterator
        With This.Items
            Do While (r < This.Count)
'          '  This always leads to a linear search but used equality
'                Do While Sequence.HasNext(Candidate)
'                    If .ElementAt(r).Equals(Candidate) Then _
'                        found = True: Exit Do
'                Loop
'
'                If Not found Then
'                    Set .ElementAt(w) = .ElementAt(r)
'                    w = w + 1
'                End If
'                found = False
                'Collection.Contains can reduce complexity depening on the collection but doesnt ensure equality compare-> can caused mismatches
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

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericOrderedList: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = GenericOrderedList: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(This): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(This): End Property
Private Property Get IGeneric_HashCode() As Long:  IGeneric_HashCode = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericOrderedList):  End Property

'@Ignore EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As Object) As Boolean: IGeneric_Equals = Me Is Other: End Function

 Private Function IGeneric_ToString() As String
    With This
        IGeneric_ToString = _
                            "Class: " & VBA.TypeName$(Me) & vbNewLine & _
                            "Capacity: " & .Items.Length & vbNewLine & _
                            "Elements: " & .Count
    
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

Private Function IGenericCollection_Contains(ByVal Element As Object) As Boolean: IGenericCollection_Contains = Not (This.Items.IndexOf(Element:=Element, Length:=This.Count) = -1): End Function

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
    
    Dim Clone As GenericOrderedList
    
    Set Clone = New GenericOrderedList
    Call Clone.AddAll(Me)
   
    Set IGenericCollection_Copy = Clone
    
End Function

Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Range: End Function

'IGenericSequence_Iterator
Private Function IGenericSequence_Iterator() As IGenericIterator: Set IGenericSequence_Iterator = Me.Range: End Function


