VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSortedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericSortedList
'*
'* Purpose: The GenericSortedList class represents a collection of elements that are sorted
'*          (based on the associated Comparer implementation) and are accessible by index.
'*          For example, if the values are of primitive types, then the SortedList is sorted in ascending order of the values.
'*          A sorted list internally maintains an array that stores the values.
'*          The capacity of a SortedList object is the number of elements the SortedList can hold.
'*          As elements are added to a SortedList, the capacity is automatically increased as required through reallocation.
'*          The capacity(is never automatically decreased) can be decreased by calling TrimToSize or by setting the Capacity property explicitly.
'*          Operations on a SortedList object tend to be slower than operations on a GenericOrderedList object because of the sorting.
'*          However, the SortedList offers better performance on search operations.
'*
'* Notes:   A SortedList object internally maintains an array to store the elements of the list
'*          It maintains ascending order (according to the associated Comparer implementation) and does allow duplicate elements.
'*          An element can be accessed by its index.
'*
'*
'*
'*======================================================================================================================

'@Folder("<T>Enumerable")
Option Explicit
'@PredeclaredId

Implements IGeneric
Implements IGenericReadOnlyList
Implements IGenericList
Implements IGenericCollection
Implements IGenericIterator

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

Private Const S_OK As Long = &H0&
Private Const DEFAULT_CAPACITY As Long = 16
Private Const SORT_ORDER As Integer = 1 'Ascending
                              
Private Type Iterator
    HasLoaded As Boolean
    Index As Long
    Count As Long
    Current As IGeneric
    Items As GenericArray
End Type

Private Type Member
    Items As GenericArray
    Comparer As IGenericComparer
    Size As Long
    Iterator As Iterator
End Type
Private myclass As Member

Public Property Get Elements() As IGenericCollection: Set Elements = Me: End Property
Public Property Get First() As Long: First = myclass.Items.LowerBound: End Property
Public Property Get Last() As Long
    
    Last = myclass.Items.LowerBound
    
    If myclass.Size > 0 Then _
        Last = myclass.Size - Last - 1
   
End Property

Public Property Get SortOrder() As SortOrder: SortOrder = SORT_ORDER: End Property
Public Property Get Comparer() As IGenericComparer: Set Comparer = myclass.Comparer: End Property

Public Property Get Capacity() As Long: Capacity = myclass.Items.Length: End Property
Public Property Let Capacity(ByVal Value As Long)
    
    If (Value < myclass.Size) Then _
        Call System.Throw(Me, "Public Property Let Capacity").ArgOutOfRange("(Value < myclass.Size)")
    
    If (myclass.Items Is Nothing) Then _
        Call System.Throw(Me, "Public Property Let Capacity").IsNothing("(myclass.Keys Is Nothing)")
    
    If Not (Value = myclass.Items.Length) Then
        If (Value > 0) Then

            Dim newItems As GenericArray: Set newItems = GenericArray.Build(Value)
            
            If (myclass.Size > 0) Then
                With myclass.Items
                    Call CopyMemory(Dst:=ByVal newItems.AdressOf(Index:=.LowerBound), Src:=ByVal .AdressOf(Index:=.LowerBound), Length:=.Size): _
                    Call ZeroMemory(Dst:=ByVal .AdressOf(Index:=.LowerBound), Length:=.Size)
                End With
            End If
           
            Set myclass.Items = newItems
           
        Else
                                                      
            Call myclass.Items.Elements.Clear
            
        End If
    End If
End Property

Public Function Build(Optional ByVal Capacity As Long, Optional ByVal Comparer As IGenericComparer) As GenericSortedList
    
    Dim newList As Member
    
    If (Capacity < 0) Then _
        Call System.Throw(Me, "Build").ArgOutOfRange("Capacity < 0")
    
    If Comparer Is Nothing Then _
        Set Comparer = IGenericComparer
    Set newList.Comparer = Comparer
    
    'Avoid null
    Set newList.Items = GenericArray
    If (Capacity > 0) Then _
        Set newList.Items = GenericArray.Build(Capacity)

    Set Build = System.CreateInstance(New GenericSortedList, VarPtr(newList), LenB(newList))
    
End Function

Public Function AsList(ByVal Comparer As IGenericComparer, ParamArray GenericValues() As Variant) As IGenericList
        
    Dim i As Long
    Dim newList As Member
    
    If IsMissing(GenericValues) Then _
        Call System.Throw(Me, "Public Function Create").NoSuchElement("ParamArray GenericValues are missing")
        
    If Comparer Is Nothing Then _
        Set Comparer = IGenericComparer
    
    newList.Size = UBound(GenericValues) + 1
    Set newList.Comparer = Comparer
    Set newList.Items = GenericArray.Build(newList.Size)
    
    With newList.Items
        For i = LBound(GenericValues) To UBound(GenericValues)
            
            If Not VarType(GenericValues(i)) = vbObject Then _
                Call System.Throw(Me, "Public Function AsList").InvalidInput("Datatype {" & TypeName$(GenericValues(i)) & "} not allowed")
            
            If (GenericValues(i) Is Nothing) Then _
                Call System.Throw(Me, "Public Function AsList").InvalidInput("Element is nothing")
            
            If Not TypeOf GenericValues(i) Is IGeneric Then _
                Call System.Throw(Me, "Public Function AsList").InvalidInput("Not TypeOf GenericValues(i) Is IGeneric")
                
            Set .ElementAt(i) = GenericValues(i)
            
        Next
                                                                              
        Call .Sort(Order:=SORT_ORDER, Comparer:=Comparer)
        
'        For i = .LowerBound + 1 To .Length - 1
'            If newList.Comparer.Compare(.ElementAt(i - 1), .ElementAt(i)) = IsEqual Then _
'                Call System.Throw(Me, "Public Function AsList").DuplicateNotAllowed(.ElementAt(i))
'        Next
            
    End With
                                                                          
    Set AsList = System.CreateInstance(New GenericSortedList, VarPtr(newList), LenB(newList))
    
End Function
                                                                                        
Private Sub Class_Terminate()
    myclass = EmptyStructure
End Sub

Public Function GetAt(ByVal Index As Long) As IGeneric

    If (Index >= myclass.Size) Then _
        Call System.Throw(Me, "Public Function GetAt").IndexOutofRange("(Index >= myclass.size)")
    
    Set GetAt = myclass.Items.ElementAt(Index)
    
End Function

Private Sub EnsureCapacity(ByVal Min As Long)
    
    Dim newCapacity As Long
    
    With myclass.Items
        If (.Length < Min) Then
            newCapacity = IIf(.Length = 0, DEFAULT_CAPACITY, .Length * 2)
            
            If (newCapacity > .MaxLength) Then _
                newCapacity = .MaxLength
            
            If (newCapacity < Min) Then _
                newCapacity = Min
        
            Me.Capacity = newCapacity
        End If
    End With
    
End Sub

'SortedList.Clear();
'SortedList.TrimExcess();
Public Sub TrimExcess()
     
    If myclass.Size < Int(CDbl(myclass.Items.Length) * 0.9) Then _
        Me.Capacity = myclass.Size

End Sub

Public Function Contains(ByVal Element As IGeneric) As Boolean: Contains = (myclass.Items.BinarySearch(Value:=Element, Order:=SORT_ORDER, Length:=myclass.Size, Comparer:=myclass.Comparer) >= 0): End Function

Public Function Distinct() As GenericSortedList
    
    Dim uniqueCount As Long
    Dim lastIndex As Long
    Dim i As Long
    
    If myclass.Size > 1 Then
        With myclass.Items
            
            lastIndex = Me.Last
            uniqueCount = .LowerBound + 1 'first element doesnt need to be shifted
    
            'Find first dup element
            Do While (uniqueCount < myclass.Size)
                If (myclass.Comparer.Compare(.ElementAt(uniqueCount - 1), .ElementAt(uniqueCount)) = IsEqual) Then _
                    Exit Do
                uniqueCount = uniqueCount + 1
            Loop
            'Shift unique elements
            For i = uniqueCount + 1 To lastIndex
                If Not (myclass.Comparer.Compare(.ElementAt(i - 1), .ElementAt(i)) = IsEqual) Then
                    Set .ElementAt(uniqueCount) = .ElementAt(i)
                    uniqueCount = uniqueCount + 1
                End If
            Next
            
            Call ClearRange(uniqueCount, lastIndex)
            myclass.Size = uniqueCount
            
        End With
    End If
    
    Set Distinct = Me
    
End Function

Public Function Remove(ByVal Element As IGeneric) As Boolean

    Dim i As Long
    i = Me.IndexOf(Element:=Element)
    
    If (i >= 0) Then _
        Call Me.RemoveAt(i): Remove = True
    
End Function

Public Sub RemoveAt(ByVal Index As Long)

    If (Index < myclass.Items.LowerBound Or Index >= myclass.Size) Then _
        Call System.Throw(Me, "Public Sub RemoveAt").ArgOutOfRange("(Index >= myclass.Size)")
   
    Set myclass.Items.ElementAt(Index) = Nothing
    
    If (Index < InterlockedDecrement(myclass.Size)) Then
        With myclass.Items
            Call CopyMemory(Dst:=ByVal .AdressOf(Index:=Index), Src:=ByVal .AdressOf(Index:=Index + 1), Length:=.SizeOf(myclass.Size - Index)): _
            Call ZeroMemory(Dst:=ByVal .AdressOf(Index:=myclass.Size), Length:=.SizeOf(1))
        End With
    End If
    
End Sub

Public Function RetainAll(ByVal Collection As IGenericCollection) As Long: RetainAll = RemoveBulk(Collection:=Collection, Check:=True): End Function
Public Function RemoveAll(ByVal Collection As IGenericCollection) As Long: RemoveAll = RemoveBulk(Collection:=Collection, Check:=False): End Function

Public Function Add(ByVal Element As IGeneric) As Boolean
       
    Dim i As Long
   
    i = myclass.Items.BinarySearch(Value:=Element, Order:=SORT_ORDER, Length:=myclass.Size, Comparer:=myclass.Comparer)
   
    If (i >= 0&) Then _
        i = Not i
        
    Call Insert(Index:=Not i, Element:=Element): Add = True
                                                                                                      
End Function

Public Function AddAll(ByVal Collection As IGenericCollection) As GenericSortedList

    Dim Element As IGeneric
    Dim Enumerable As IGenericIterator
    Dim sorted As Boolean
    Dim sortedList As GenericSortedList
    Dim newElements As GenericArray
            
    If (Collection Is Nothing) Then _
        Call System.Throw(Me, "Public Sub AddAll").IsNothing("(Collection Is Nothing)")
   
    If (Collection.Count > 0) Then
        Set newElements = Collection.ToArray
        
        If TypeOf Collection Is GenericSortedList Then
            Set sortedList = Collection
            If (IGenericComparer.Equals(myclass.Comparer, sortedList.Comparer)) Then _
                sorted = True
        End If
    
        If Not sorted Then _
            Call newElements.Sort(Order:=SORT_ORDER, Comparer:=myclass.Comparer)
        
        If myclass.Size > 0 Then
            Set myclass.Items = Merge(SortedA:=myclass.Items, LengthA:=myclass.Size, SortedB:=newElements, LengthB:=newElements.Length, Comparer:=myclass.Comparer, Order:=SORT_ORDER)
        Else
            Set myclass.Items = newElements
        End If
        
        myclass.Size = myclass.Items.Length

    Else
'        Do it the hard way
        Set Enumerable = Collection
        With Enumerable
            Do While .HasNext(Element): Call Me.Add(Element:=Element): Loop
        End With
        
    End If
    
    Set AddAll = Me
    
End Function

Private Sub Insert(ByVal Index As Long, ByVal Element As IGeneric)

    If (myclass.Size = myclass.Items.Length) Then _
        Call EnsureCapacity(myclass.Size + 1)
    
    With myclass.Items
        If (Index < myclass.Size) Then _
            Call CopyMemory(Dst:=ByVal .AdressOf(Index:=Index + 1), Src:=ByVal .AdressOf(Index:=Index), Length:=.SizeOf(myclass.Size - Index)): _
            Call ZeroMemory(Dst:=ByVal .AdressOf(Index:=Index), Length:=.SizeOf(1))
    
        Set .ElementAt(Index) = Element
    End With
    
    myclass.Size = myclass.Size + 1

End Sub

'Returns the index of the entry in this sorted list. The
'entry is located through a binary search, and thus the average execution
'time of this method is proportional to Log2(Size), where
'Size is the Size of this sorted list. The returned entry is -1 if
'the given entry does not occur in this sorted list. Null is an valid
'entry value.
Public Function IndexOf(ByVal Element As IGeneric) As Long
     
    Dim ret As Long
    ret = myclass.Items.BinarySearch(Value:=Element, Order:=SORT_ORDER, Length:=myclass.Size, Comparer:=myclass.Comparer)

    IndexOf = IIf(ret >= 0, ret, -1)
    
End Function

'Pops an item from the top of the stack.  If the stack is empty, Pop
'throws an InvalidOperationException.
Public Function Pop() As IGeneric
    
    Dim i As Long
    
    With Me
        i = .Last
        Set Pop = .GetAt(i)
        Call .RemoveAt(i)
    End With
    
End Function
'Pushes an item to the top of the stack.
Public Function Push(ByVal Element As IGeneric) As GenericOrderedList
    Call Me.Add(Element:=Element)
    Set Push = Me
End Function
'Returns the top object on the stack without removing it.  If the stack
'is empty, Peek throws an InvalidOperationException.
Public Function Peek() As IGeneric
    With Me
        Set Peek = .GetAt(.Last)
    End With
End Function
                                                                                            
Public Sub CopyTo( _
                    ByVal GArray As GenericArray, _
                    ByVal ArrayIndex As Long, _
                    Optional ByVal Index As Long = 0, _
                    Optional ByVal Count As Long = 0)

    If (Count = 0) Then _
        Count = myclass.Size - Index
        
    If (myclass.Size - Index < Count) Then _
        Call System.Throw(Me, "Public Sub CopyTo").ArgOutOfRange("(myclass.Size - Index < Count)")
    
    Call myclass.Items.CopyTo(GArray:=GArray, ArrayIndex:=ArrayIndex, Index:=Index, Count:=Count)

End Sub
                                                                                                    
Private Function RemoveBulk(ByVal Collection As IGenericCollection, ByVal Check As Boolean) As Long
    
    Dim Candidate As IGeneric
    Dim result As Long
    Dim r As Long
    Dim w As Long
    
    If (Collection Is Nothing) Then _
        Call System.Throw(Me, "Public Function RemoveAll").IsNothing("(Collection Is Nothing)")
    
    If (Collection.Count > 0) Then
        With myclass.Items
            Do While (r < myclass.Size)
                Set Candidate = .ElementAt(r)
                If (Collection.Contains(Candidate) = Check) Then _
                    Set .ElementAt(w) = Candidate: w = w + 1
              
                r = r + 1
            Loop
            
            If Not (r = myclass.Size) Then _
                Call System.Throw(Me, "Public Sub RemoveAll").InvalidOperation
            
            Call ClearRange(w, myclass.Size - 1)
            result = myclass.Size - w
           
        End With
    End If
    
    myclass.Size = myclass.Size - result
    RemoveBulk = result
    
End Function

Private Sub ClearRange(ByVal BeginAt As Long, ByVal EndAt As Long)
    Dim i As Long
    With myclass.Items
        For i = BeginAt To EndAt: Set .ElementAt(i) = Nothing: Next
    End With
End Sub

Private Function Merge( _
                        ByVal SortedA As GenericArray, _
                        ByVal LengthA As Long, _
                        ByVal SortedB As GenericArray, _
                        ByVal LengthB As Long, _
                        Optional ByVal Comparer As IGenericComparer, _
                        Optional ByVal Order As SortOrder = SORT_ORDER) As GenericArray
    
    Dim i As Long, j As Long, k As Long
    Dim A As IGeneric, B As IGeneric
    
    Dim Merged As GenericArray
    Set Merged = GenericArray.Build(LengthA + LengthB)
    
    If Comparer Is Nothing Then _
        Set Comparer = IGenericComparer
        
    With Merged
    
        Do While (i < LengthA And j < LengthB)
            Set A = SortedA.ElementAt(i)
            Set B = SortedB.ElementAt(j)

            If (Comparer.Compare(A, B) * Order = islower) Then
                Set .ElementAt(k) = A: i = i + 1
            Else
                Set .ElementAt(k) = B: j = j + 1
            End If

            k = k + 1
        Loop

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
    
    Set Merge = Merged

End Function

Private Sub LoadIterator(ByRef Iterator As Iterator, ByVal Items As GenericArray, ByVal Count As Long)
    
    If (Items Is Nothing) Then _
        Call System.Throw(Me, "Private Sub LoadIterator").InvalidInput("(Items Is Nothing)")
    
    If (Items.IsEmpty) Then _
        Call System.Throw(Me, "Private Sub LoadIterator").InvalidInput("(Items.IsEmpty)")
        
    If (Count > Items.Length) Then _
        Call System.Throw(Me, "Private Sub LoadIterator").InvalidInput("(Count >= Items.Length)")
    
    Set Iterator.Items = Items
    Iterator.Index = Items.LowerBound
    Iterator.Count = Count
    Iterator.HasLoaded = True
    
End Sub

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericSortedList: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GenericSortedList: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericSortedList):  End Property
Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = Me.Elements.Copy: End Function

Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    If Other Is Nothing Then _
        IGeneric_Equals = False: Exit Function
    
    IGeneric_Equals = (Me Is Other)
End Function

Private Function IGeneric_ToString() As String
    With myclass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Capacity: " & .Items.Length & vbNewLine & _
                            "Elements: " & .Size & vbNewLine & _
                            "Sortorder: " & IIf(SORT_ORDER = 1, "Ascending", "Descending")
    End With
End Function

'IGenericList
Private Property Get IGenericList_Elements() As IGenericCollection: Set IGenericList_Elements = Me: End Property
Private Function IGenericList_Add(ByVal Element As IGeneric) As Boolean: IGenericList_Add = Me.Add(Element:=Element): End Function
Private Function IGenericList_AddAll(ByVal Collection As IGenericCollection) As IGenericList: Set IGenericList_AddAll = Me.AddAll(Collection:=Collection): End Function
Private Sub IGenericList_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long, Optional ByVal Index As Long = 0&, Optional ByVal Count As Long = 0&): Call Me.CopyTo(GArray:=GArray, ArrayIndex:=ArrayIndex, Index:=Index, Count:=Count): End Sub
Private Function IGenericList_GetAt(ByVal Index As Long) As IGeneric: Set IGenericList_GetAt = Me.GetAt(Index:=Index): End Function
Private Function IGenericList_IndexOf(ByVal Element As IGeneric) As Long: IGenericList_IndexOf = Me.IndexOf(Element:=Element): End Function
Private Function IGenericList_Pop() As IGeneric: Set IGenericList_Pop = Me.Pop: End Function
Private Function IGenericList_Remove(ByVal Element As IGeneric) As Boolean: IGenericList_Remove = Me.Remove(Element:=Element): End Function
Private Function IGenericList_RemoveAll(ByVal Collection As IGenericCollection) As Long: IGenericList_RemoveAll = Me.RemoveAll(Collection:=Collection): End Function
Private Function IGenericList_RetainAll(ByVal Collection As IGenericCollection) As Long: IGenericList_RetainAll = RemoveBulk(Collection:=Collection, Check:=True): End Function
Private Sub IGenericList_RemoveAt(ByVal Index As Long): Call Me.RemoveAt(Index:=Index): End Sub

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_Elements() As IGenericCollection: Set IGenericReadOnlyList_Elements = Me: End Property
Private Property Get IGenericReadOnlyList_First() As Long: IGenericReadOnlyList_First = Me.First: End Property
Private Property Get IGenericReadOnlyList_Last() As Long: IGenericReadOnlyList_Last = Me.Last: End Property

Private Function IGenericReadOnlyList_GetAt(ByVal Index As Long) As IGeneric
    
    If (Index >= myclass.Size) Then _
        Call System.Throw(Me, "IGenericReadOnlyList_GetAt").IndexOutofRange("(Index >= myclass.size)")
    
    Set IGenericReadOnlyList_GetAt = myclass.Items.ElementAt(Index)
    
End Function
Private Function IGenericReadOnlyList_IndexOf(ByVal Element As IGeneric) As Long: IGenericReadOnlyList_IndexOf = Me.IndexOf(Element:=Element): End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = myclass.Size: End Property
'ToArray returns a new Object array containing the contents of the List.
'This requires copying the List, which is an O(n) operation.
Private Function IGenericCollection_ToArray() As GenericArray

    Dim Clone As GenericArray
    Set Clone = GenericArray.Build(myclass.Size)
    Call Me.CopyTo(Clone, Clone.LowerBound, Me.First, myclass.Size)
    
    Set IGenericCollection_ToArray = Clone

End Function

Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = myclass.Items.BinarySearch(Value:=Item, Order:=SORT_ORDER, Length:=myclass.Size, Comparer:=myclass.Comparer) >= 0: End Function

Private Function IGenericCollection_ContainsAll(ByVal Collection As IGenericCollection) As Boolean
    
    Dim Candidate As IGeneric
    Dim Enumerable As IGenericIterator
    
    If (Collection Is Nothing) Then _
        Call System.Throw(Me, "Private Function IGenericCollection_ContainsAll").IsNothing("(Collection Is Nothing)")
    
    If Not (Collection Is Me) Then
    
        If (Collection.Count > 0) Then
            Set Enumerable = Collection.Iterator
        Else
            Set Enumerable = Collection
        End If
        
        With Enumerable
            Do While .HasNext(Candidate)
                If (myclass.Items.BinarySearch(Value:=Candidate, Order:=SORT_ORDER, Length:=myclass.Size, Comparer:=myclass.Comparer) < 0) Then _
                    Exit Function
            Loop
        End With
        
    End If

    IGenericCollection_ContainsAll = True

End Function

Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call Me.CopyTo(Index:=myclass.Items.LowerBound, GArray:=GArray, ArrayIndex:=ArrayIndex, Count:=myclass.Size): End Sub

Private Sub IGenericCollection_Clear()
    
    If Not (myclass.Items Is Nothing) Then _
        Call myclass.Items.Elements.Clear
    myclass.Size = 0

End Sub

Private Function IGenericCollection_Copy() As IGenericCollection
    
    Dim C As Member
    
    Set C.Comparer = myclass.Comparer
    Set C.Items = myclass.Items.Elements.Copy
    C.Size = myclass.Size
 
    Set IGenericCollection_Copy = System.CreateInstance(New GenericSortedList, VarPtr(C), LenB(C))
    
End Function

Private Function IGenericCollection_Iterator() As IGenericIterator

    Dim C As Member
    
    Call LoadIterator(C.Iterator, myclass.Items, myclass.Size)
    Set IGenericCollection_Iterator = System.CreateInstance(New GenericSortedList, VarPtr(C), LenB(C))

End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.Current: End Property
Private Function IGenericIterator_HasNext(Optional ByRef Out As IGeneric) As Boolean
    
    With myclass.Iterator
        If .Index < .Count Then
            Set .Current = .Items.ElementAt(.Index)
            .Index = .Index + 1
            Set Out = .Current
            IGenericIterator_HasNext = True
            Exit Function
        End If
        
        If .HasLoaded = False Then _
            Call System.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator

        Set .Current = Nothing
     
    End With
    
End Function

Private Sub IGenericIterator_Reset()
    With myclass.Iterator
        If .HasLoaded = False Then _
            Call System.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
        
        Call LoadIterator(myclass.Iterator, .Items, .Count)
    End With
End Sub


