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
'*          An element cannot be null and must be unique.
'*          For example, if the values are of primitive types, then the SortedList is sorted in ascending order of the values.
'*          A sorted list internally maintains an array that stores the values.
'*          The capacity of a SortedList object is the number of elements the SortedList can hold.
'*          As elements are added to a SortedList, the capacity is automatically increased as required through reallocation.
'*          The capacity(is never automatically decreased) can be decreased by calling TrimToSize or by setting the Capacity property explicitly.
'*          Operations on a SortedList object tend to be slower than operations on a GenericOrderedList object because of the sorting.
'*          However, the SortedList offers better performance on search operations.
'*
'* Notes:   A SortedList object internally maintains an array to store the elements of the list
'*          Elements of a SortedList object are sorted according to the associated Comparer implementation
'*          By default, this collection sort the elements in ascending order.
'*          A element must be unique and cannot be null.
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
Private Const SORT_ORDER As Long = 1 'Ascending
                              
Private Type Iterator
    HasLoaded As Boolean
    Index As Long
    Count As Long
    current As IGeneric
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

Public Property Get ElementAt(ByVal Index As Long) As IGeneric
Attribute ElementAt.VB_UserMemId = 0

    If (Index >= myclass.Size) Then _
        Call System.Throw(Me, "Public Property Get ElementAt").IndexOutofRange("(Index >= myclass.size)")
    
    Set ElementAt = myclass.Items.ElementAt(Index)
    
End Property

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

Public Function AsList(ByVal Comparer As IGenericComparer, ParamArray GenericValues() As Variant) As GenericSortedList
        
    Dim i As Long
    Dim newList As Member
    
    If IsMissing(GenericValues) Then _
        Call System.Throw(Me, "Public Function Create").Instantiation("ParamArray GenericValues are missing")
        
    If Comparer Is Nothing Then _
        Set Comparer = IGenericComparer
    
    newList.Size = UBound(GenericValues) + 1
    Set newList.Comparer = Comparer
    Set newList.Items = GenericArray.Build(newList.Size)
    
    With newList.Items
        For i = LBound(GenericValues) To UBound(GenericValues)
            
            If Not VarType(GenericValues(i)) = vbObject Then _
                Call System.Throw(Me, "Public Function Create").InvalidInput("Datatype {" & TypeName$(GenericValues(i)) & "} not allowed")
            
            If (GenericValues(i) Is Nothing) Then _
                Call System.Throw(Me, "Public Function Create").InvalidInput("Element is nothing")
            
            If Not TypeOf GenericValues(i) Is IGeneric Then _
                Call System.Throw(Me, "Public Function Create").InvalidInput("Not TypeOf GenericValues(i) Is IGeneric")
                
            Set .ElementAt(i) = GenericValues(i)
            
        Next
                                                                              
        Call .Sort(Order:=SORT_ORDER, Index:=.LowerBound, Length:=.Length, Comparer:=Comparer)
        
        For i = .LowerBound + 1 To .Length - 1
            If Comparer.Compare(.ElementAt(i - 1), .ElementAt(i)) = IsEqual Then _
                Call System.Throw(Me, "Public Function Create").DuplicateNotAllowed(.ElementAt(i))
        Next
    End With
                                                                          
    Set AsList = System.CreateInstance(New GenericSortedList, VarPtr(newList), LenB(newList))
    
End Function
                                                                                        
Private Sub Class_Terminate()
    myclass = EmptyStructure
End Sub

Private Sub EnsureCapacity(ByVal min As Long)
    
    Dim newCapacity As Long
    
    With myclass.Items
        If (.Length < min) Then
            newCapacity = IIf(.Length = 0, DEFAULT_CAPACITY, .Length * 2)
            
            If (newCapacity > .MaxLength) Then _
                newCapacity = .MaxLength
            
            If (newCapacity < min) Then _
                newCapacity = min
        
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

Public Function Contains(ByVal Element As IGeneric) As Long: Contains = Me.IndexOf(Element:=Element) > -1: End Function

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

Public Function RemoveAll(ByVal Collection As IGenericCollection) As Long

    Dim Candidate As IGeneric
    Dim result As Long
    Dim i As Long

    If (Collection Is Nothing) Then _
        Call System.Throw(Me, "Public Sub RemoveAll").IsNothing("(Collection Is Nothing)")
    
    If (Collection.Count > 0) Then
        With Collection.Iterator
            Do While .HasNext(Candidate)
                i = Me.IndexOf(Element:=Candidate)
                If i >= 0 Then _
                    Call Me.RemoveAt(i): result = result + 1
                 
            Loop
        End With
    End If
    
    myclass.Size = myclass.Size - result
    RemoveAll = result

End Function

Public Function Add(ByVal Element As IGeneric) As Boolean
       
    Dim i As Long
    Dim LB As Long
 
    LB = myclass.Items.LowerBound
    i = myclass.Items.BinarySearch(Value:=Element, Order:=SORT_ORDER, Index:=LB, Length:=myclass.Size, Comparer:=myclass.Comparer)
   
    If (i < LB) Then _
        Call Insert(Index:=Not i, Element:=Element): Add = True
                                                                                                      
End Function

Public Sub AddAll(ByVal Collection As IGenericCollection)

    Dim Element As IGeneric, ElementBefore As IGeneric
    Dim Enumerable As IGenericIterator
    Dim Count As Long, i As Long
  
    If (Collection Is Nothing) Then _
        Call System.Throw(Me, "Public Sub AddAll").IsNothing("(Collection Is Nothing)")
    
    If (Collection Is Me) Then _
        Exit Sub
        
    Count = Collection.Count
    If (Count > 0) Then
    
        With Collection.ToArray.Sort(Order:=SORT_ORDER, Index:=GenericArray.LowerBound, Length:=Count, Comparer:=myclass.Comparer)
            For i = .LowerBound To .Length - 1
                Set Element = .ElementAt(i)
                If Not myclass.Comparer.Compare(Element, ElementBefore) = IsEqual Then _
                    Call Me.Add(Element:=Element)
                Set ElementBefore = Element
            Next
        End With
        
    Else
    
        Set Enumerable = Collection 'Try
        With Enumerable
            Do While .HasNext(Element): Call Me.Add(Element:=Element): Loop
        End With
        
    End If
    
End Sub

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
        Set Pop = .ElementAt(i)
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
        Set Peek = .ElementAt(.Last)
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
    
    Call GenericArray.CopyArrays(SrcArray:=myclass.Items, SrcIndex:=Index, DstArray:=GArray, DstIndex:=ArrayIndex, Elements:=Count)

End Sub
                                                                                                    
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

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_Elements() As IGenericCollection: Set IGenericReadOnlyList_Elements = Me: End Property
Private Property Get IGenericReadOnlyList_First() As Long: IGenericReadOnlyList_First = Me.First: End Property
Private Property Get IGenericReadOnlyList_Last() As Long: IGenericReadOnlyList_Last = Me.Last: End Property

Private Property Get IGenericReadOnlyList_ElementAt(ByVal Index As Long) As IGeneric
    
    If (Index >= myclass.Size) Then _
        Call System.Throw(Me, "Private Property Get IGenericReadOnlyList_ElementAt").IndexOutofRange("(Index >= myclass.size)")
    
    Set IGenericReadOnlyList_ElementAt = myclass.Items.ElementAt(Index)
    
End Property
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
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call Me.CopyTo(Index:=myclass.Items.LowerBound, GArray:=GArray, ArrayIndex:=ArrayIndex, Count:=myclass.Size): End Sub

Private Sub IGenericCollection_Clear()
    
    If Not (myclass.Items Is Nothing) Then _
        Call myclass.Items.Elements.Clear
    myclass.Size = 0

End Sub

Private Function IGenericCollection_Copy() As IGenericCollection
    
    Dim c As Member
    
    Set c.Comparer = myclass.Comparer
    Set c.Items = myclass.Items.Elements.Copy
    c.Size = myclass.Size
 
    Set IGenericCollection_Copy = System.CreateInstance(New GenericSortedList, VarPtr(c), LenB(c))
    
End Function

Private Function IGenericCollection_Iterator() As IGenericIterator

    Dim c As Member
    
    Call LoadIterator(c.Iterator, myclass.Items, myclass.Size)
    Set IGenericCollection_Iterator = System.CreateInstance(New GenericSortedList, VarPtr(c), LenB(c))

End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    With myclass.Iterator
        If .Index < .Count Then
            Set .current = .Items.ElementAt(.Index)
            .Index = .Index + 1
            Set Out = .current
            IGenericIterator_HasNext = True
            Exit Function
        End If
        
        If .HasLoaded = False Then _
            Call System.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator

        Set .current = Nothing
        Set Out = .current
     
    End With
    
End Function

Private Sub IGenericIterator_Reset()
    With myclass.Iterator
        If .HasLoaded = False Then _
            Call System.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
        
        Call LoadIterator(myclass.Iterator, .Items, .Count)
    End With
End Sub


