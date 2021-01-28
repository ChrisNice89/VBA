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
'* Purpose: The GenericSortedList class represents a collection of key-and-value pairs that are sorted by the keys
'*          (based on the associated Comparer implementation) and are accessible by key and by index.
'*          A key cannot be null and must be unique, but a value can be.
'*          For example, if the keys are of primitive types, then the SortedList is sorted in ascending order of the keys.
'*          A sorted list internally maintains two arrays that store the keys and values of the entries.
'*          The capacity of a SortedList object is the number of elements the SortedList can hold.
'*          As elements are added to a SortedList, the capacity is automatically increased as required through reallocation.
'*          The capacity(is never automatically decreased) can be decreased by calling TrimToSize or by setting the Capacity property explicitly.
'*          Operations on a SortedList object tend to be slower than operations on a Hashtable object because of the sorting.
'*          However, the SortedList offers more flexibility by allowing access to the values either through
'*          the associated keys or through the indexes.
'*
'* Notes:   A SortedList object internally maintains two arrays to store the elements of the list,
'*          i.e, one array for the keys and another array for the associated values.
'*          Elements of a SortedList object are sorted by the keys according to the IGeneric Compare implementation
'*          provided by the keys themselves.
'*          By default, this collection sort the key/value pairs in ascending order.
'*          A key must be unique and cannot be null.
'*          A value can be null or duplicate.
'*          An element can be accessed by its key or by its index.

'*          The SortedList class provides a convenient way to create a sorted
'*          copy of another dictionary, such as a GenericMap. For example:
'*
'*          h = GenericMap.Build();
'*          h.Add(...);
'*          h.Add(...);
'*
'*          sl = GenericSortedList.BuildFrom(Dictionary:=h);
'*
'*          The last line above creates a sorted list that contains a copy of the keys
'*          and values stored in the hashtable. In this particular example, the keys
'*          will be ordered according to the Comparer implementation, which they
'*          all must implement.
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

Private Type Iterator
    StartIndex As Long
    EndIndex As Long
    Count As Long
    Index As Long
    HasLoaded As Boolean
    Current As IGeneric
    Items As GenericArray
End Type

Private Type Member
    Items As GenericArray
    Comparer As IGenericComparer
    Size As Long
    SortOrder As SortOrder
    Iterator As Iterator
End Type
Private myclass As Member

Public Property Get First() As Long: First = myclass.Items.LowerBound: End Property
Public Property Get Last() As Long
    
    Last = myclass.Items.LowerBound
    
    If myclass.Size > 0 Then _
        Last = myclass.Size - Last - 1
   
End Property

Public Property Get SortOrder() As SortOrder: SortOrder = myclass.SortOrder: End Property
Public Property Get Count() As Long: Count = myclass.Size: End Property
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
                    Call CopyMemory(Dst:=ByVal newKeys.AdressOf(Index:=.LowerBound), Src:=ByVal .AdressOf(Index:=.LowerBound), Length:=.Size): _
                    Call ZeroMemory(Dst:=ByVal .AdressOf(Index:=.LowerBound), Length:=.Size)
                End With
            End If
           
            Set myclass.Items = newItems
           
        Else
                                                      
            Call myclass.Items.Clear
            
        End If
    End If
End Property

Public Function Build(Optional ByVal Capacity As Long = DEFAULT_CAPACITY, Optional ByVal Comparer As IGenericComparer) As GenericSortedList
    
    Dim newList As Member
    
    If (Capacity < 0) Then _
        Call System.Throw(Me, "Build").ArgOutOfRange("Capacity < 0")
    
    If Comparer Is Nothing Then _
        Set Comparer = IGenericComparer
    
    newList.SortOrder = Ascending
    Set newList.Comparer = Comparer
    Set newList.Items = GenericArray.Build(Capacity)

    Set Build = System.CreateInstance(New GenericSortedList, VarPtr(newList), LenB(newList))
    
End Function

Public Function BuildFrom(ByVal Dictionary As IGenericDictionary, Optional ByVal Comparer As IGenericComparer) As GenericSortedList
    
    Dim newList As Member
    
    If (Dictionary Is Nothing) Then _
        Call System.Throw(Me, "Public Function BuildFrom").IsNothing("(Dictionary Is Nothing)")
    
    newList.SortOrder = Ascending
    
    If Comparer Is Nothing Then _
        Set Comparer = IGenericValue.Comparer
    
    Set newList.Comparer = Comparer
    Set newList.Keys = GenericArray.Build(Dictionary.Count)
    Set newList.Values = GenericArray.Build(Dictionary.Count)
    
    Call Dictionary.CopyOf(t:=KeyData, GArray:=newList.Keys, ArrayIndex:=newList.Keys.LowerBound)
    Call Dictionary.CopyOf(t:=ValueData, GArray:=newList.Values, ArrayIndex:=newList.Values.LowerBound)
    
    Call GenericArray.SortGenericArrays(Keys:=newList.Keys, Items:=newList.Values, Index:=newList.Keys.LowerBound, Length:=newList.Keys.Length, Comparer:=newList.Comparer, Order:=newList.SortOrder)
    newList.Size = Dictionary.Count
    
    Set BuildFrom = System.CreateInstance(New GenericSortedList, VarPtr(newList), LenB(newList))
    
End Function

Public Function Create(ByVal Comparer As IGenericComparer, ParamArray GenericValues() As Variant) As GenericSortedSet
        
    Dim i As Long
    Dim Values As GenericArray
    Dim newList As Member
    
    If IsMissing(GenericValues) Then _
        Call System.Throw(Me, "Public Function Create").Instantiation("ParamArray GenericValues are missing")
        
    If Comparer Is Nothing Then _
        Set Comparer = IGenericComparer
    
    Set newList.Comparer = Comparer
    newList.SortOrder = Ascending  
                                                                                    
    Set newList.Values = GenericArray.Build(UBound(GenericValues) + 1)
    
    With newList.Values
        For i = LBound(GenericValues) To UBound(GenericValues)
            
            If Not VarType(GenericValues(i)) = vbObject Then _
                Call System.Throw(Me, "Public Function Create").InvalidInput("Datatype {" & TypeName$(GenericValues(i)) & "} not allowed")
            
            If (GenericValues(i) Is Nothing) Then _
                Call System.Throw(Me, "Public Function Create").InvalidInput("Element is nothing")
            
            If Not TypeOf GenericValues(i) Is IGeneric Then _
                Call System.Throw(Me, "Public Function Create").InvalidInput("Not TypeOf GenericValues(i) Is IGeneric")
                
            Set .ElementAt(i) = GenericValues(i)
          
        Next
        
        Call .Sort(Order:=newList.SortOrder, Index:=.LowerBound, Length:=.Length, Comparer:=newList.Comparer)  
                                                                                            
    End With
                                                                           
    Set Create =   System.CreateInstance(New GenericSortedList, VarPtr(newList), LenB(newList))
    
End Function
                                                                                        
Private Sub Class_Terminate()
    Call IGeneric_Dispose
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
     
    If myclass.Size < Int(CDbl(myclass.items.Length) * 0.9) Then _
        Me.Capacity = myclass.Size

End Sub

Public Function Contains(ByVal Value As IGeneric) As Long: Contains = Me.IndexOf(Value) > -1: End Function

Public Function Remove(ByVal Value As IGeneric) As Boolean

    Dim i As Long
    i = Me.IndexOf(Value)
    
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

Public Function Add(ByVal Value As IGeneric) As Long
   
    If (Value Is Nothing) Then _
        Call System.Throw(Me, "Public Function Add").IsNothing("(Value Is Nothing)")
       
    Dim i As Long
    i = Not myclass.Items.BinarySearch(Value:=Key, Order:=myclass.SortOrder, Index:=myclass.Items.LowerBound, Length:=myclass.Size, Comparer:=myclass.Comparer)
   
    If (i >= 0) Then _
        Call System.Throw(Me, "Public Function Add").InvalidInput("Value duplicated")

    Call Insert(Index:=Not i, Value:=Value)

End Function

Private Sub Insert(ByVal Index As Long,  ByVal Value As IGeneric)

    If (myclass.Size = myclass.Items.Length) Then _
        Call EnsureCapacity(myclass.Size + 1)
    
    With myclass.Items                                                                                  
        If (Index < myclass.Size) Then _
            Call CopyMemory(Dst:=ByVal .AdressOf(Index:=Index + 1), Src:=ByVal .AdressOf(Index:=Index), Length:=.SizeOf(myclass.Size - Index)): _
            Call ZeroMemory(Dst:=ByVal .AdressOf(Index:=Index), Length:=.SizeOf(1))
    
        Set .ElementAt(Index) = Value
    end with                                                                                      
    
    myclass.Size = myclass.Size + 1

End Sub

'Returns the index of the entry with a given key in this sorted list. The
'key is located through a binary search, and thus the average execution
'time of this method is proportional to Log2(Size), where
'Size is the Size of this sorted list. The returned value is -1 if
'the given key does not occur in this sorted list. Null is an invalid
'key value.
Public Function IndexOf(ByVal Value As IGeneric) As Long
     
    If (Value Is Nothing) Then _
        Call System.Throw(Me, "Public Function IndexOf").IsNothing("(Value Is Nothing)")
    
    Dim ret As Long
    ret = Not myclass.Items.BinarySearch(Value:=Value, Index:=myclass.Itemss.LowerBound, Length:=myclass.Size, Order:=myclass.SortOrder, Comparer:=myclass.Comparer)

    IndexOf = IIf(ret >= 0, ret, -1)
    
End Function

Public Sub Clear():
    Call myclass.Items.Clear
    myclass.Size = 0
End Sub

Public Function Iterator(Optional ByVal Index As Long, Optional ByVal Count As Long) As IGenericIterator
    
    Dim c As Member
    
    If Count = 0 Then _
        Count = myclass.Size
    
    If (Index < myclass.Items.LowerBound) Then _
        Call System.Throw(Me, "Private Sub LoadIterator").IndexOutofRange("(Index < myclass.Items.LowerBound)")

    If (Count < 0) Or (Count > myclass.Size - Index ) Then _
        Call System.Throw(Me, "Public Function Iterator").ArgOutOfRange("(Count < 0) Or (Count > myclass.Size - Index )")
        
    Call LoadIterator(c.Iterator, myclass.Items, Index, Count)
    Set Iterator = System.CreateInstance(New GenericSortedList, VarPtr(c), LenB(c))

End Function

Public Sub CopyTo( _
                    ByVal Index As Long, _
                    ByVal GArray As GenericArray, _
                    ByVal ArrayIndex As Long, _
                    Optional ByVal Count As Long = 0)

    If (Count = 0) Then _
        Count = myclass.Size
        
    If (myclass.Size - Index < Count) Then _
        Call System.Throw(Me, "Public Sub CopyTo").ArgOutOfRange("(myclass.Size - Index<count)")
    
    Call GenericArray.CopyArrays(SrcArray:=myclass.Items, SrcIndex:=Index, DstArray:=GArray, DstIndex:=ArrayIndex, Elements:=Count)

End Sub   
                                                                                                  
'ToArray returns a new Object array containing the contents of the List.
'This requires copying the List, which is an O(n) operation.
Public Function ToArray() As GenericArray
   
    Set ToArray = GenericArray.Build(myclass.Size)
    Call Me.CopyTo(Index:=myclass.Items.LowerBound, GArray:=ToArray, ArrayIndex:=ToArray.LowerBound)

End Function
                                                                                                    
Private Sub LoadIterator( _
                            ByRef Iterator As Iterator, _
                            ByVal Items As GenericArray, _
                            ByVal StartIndex As Long, _
                            ByVal Count As Long)

    Set Iterator.Items = Items
    Iterator.StartIndex = StartIndex
    Iterator.Count = Count
    Iterator.Index = StartIndex - 1
    Iterator.EndIndex = StartIndex + Count 
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

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member

    Set c.Items = myclass.Items.Copy
    c.Size = myclass.Size
    c.SortOrder=myclass.sortorder                                                                                                                
 
    Set IGeneric_Clone = System.CreateInstance(New GenericSortedList, VarPtr(c), LenB(c))
   
End Function

Private Sub IGeneric_Dispose()
    myclass = EmptyStructure
End Sub
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
                            "Sortorder: " & IIf(.SortOrder = Ascending, "Ascending", "Descending")
    End With
End Function

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_First() As Long: IGenericReadOnlyList_First = Me.First: End Property
Private Property Get IGenericReadOnlyList_Last() As Long: IGenericReadOnlyList_Last = Me.Last: End Property
Private Property Get IGenericReadOnlyList_Count() As Long: IGenericReadOnlyList_Count = myclass.Size: End Property
Private Property Get IGenericReadOnlyList_ElementAt(ByVal Index As Long) As IGeneric
    
    If (Index >= myclass.Size) Then _
        Call System.Throw(Me, "Private Property Get IGenericReadOnlyList_ElementAt").IndexOutofRange("(Index >= myclass.size)")
    
    Set IGenericReadOnlyList_ElementAt = myclass.Items.ElementAt(Index)
    
End Property
Private Function IGenericReadOnlyList_IndexOf(ByVal Item As IGeneric) As Long: IGenericReadOnlyList_IndexOf = Me.IndexOf(Value:=Item): End Function
Private Sub IGenericReadOnlyList_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call Me.CopyTo(Index:=myclass.Items.LowerBound, GArray:=GArray, ArrayIndex:=ArrayIndex, Count:=myclass.Size): End Sub
Private Function IGenericReadOnlyList_Iterator() As IGenericIterator: Set IGenericReadOnlyList_Iterator = Me.Iterator: End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = me.toarray: End Function
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.contains(Value:=Item): End Function
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call Me.CopyTo(Index:=myclass.Items.LowerBound, GArray:=GArray, ArrayIndex:=ArrayIndex, Count:=myclass.Size): End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Iterator: End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.Current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    With myclass.Iterator
    
        If .Index < .EndIndex Then
            Set .Current = .Items.ElementAt(InterlockedIncrement(.Index))
            Set Out= .Current
            IGenericIterator_HasNext=True
            exit function
        end if
        
        If .HasLoaded = False Then 
            Call System.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator

        set .current=nothing
        Set Out= .Current
     
    End With
    
End Function

Private Sub IGenericIterator_Reset()
    With myclass.Iterator
        If .HasLoaded = False Then _
            Call System.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
        
        Call LoadIterator(myclass.Iterator, .Items, .StartIndex, .Count)
    End With
End Sub


