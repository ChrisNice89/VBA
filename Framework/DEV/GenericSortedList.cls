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
'*          (based on the associated IGeneric CompareTo implementation) and are accessible by key and by index.
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
'*          will be ordered according to the IGeneric Compare implementation, which they
'*          all must implement.
'*
'*
'*
'*======================================================================================================================

'@Folder("<T>Enumerable")
Option Explicit
'@PredeclaredId

Implements IGeneric
Implements IGenericDictionary
Implements IGenericReadOnlyList
Implements IGenericCollection
Implements IGenericIterator

#If Win64 Then
    Private Const POINTERSIZE As LongPtr = 8
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst As Any, ByRef src As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dst As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef dst As Any, ByVal Length As LongPtr, ByVal Fill As Byte)
    Private Declare PtrSafe Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As LongPtr, ByRef dst As LongPtr, ByVal src As LongPtr) As Long
    Private Declare PtrSafe Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Length As Long, ByRef dst As LongPtr, ByVal src As LongPtr) As Long
    Private Declare PtrSafe Function SafeArrayGetElement Lib "oleaut32.dll" (ByVal psa As Any, ByRef rgIndices As Long, ByRef pv As Any) As LongPtr
    Private Declare PtrSafe Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As LongPtr
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
    Private Const POINTERSIZE As Long = 4
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst As Any, ByRef src As Any, ByVal Length As Long)
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dst As Any, ByVal Length As Long)
    Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef dst As Any, ByVal Length As Long, ByVal Fill As Byte)
    Private Declare Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As Long, ByRef dst As Long, ByVal src As Long) As Long
    Private Declare Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Length As Long, ByRef dst As Long, ByVal src As Long) As Long
    Private Declare PtrSafe Function SafeArrayGetElement Lib "oleaut32.dll" (ByVal psa As Any, ByRef rgIndices As Long, ByRef pv As Any) As LongPtr
    Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

Private Const S_OK As Long = &H0&
Private Const DEFAULT_CAPACITY As Long = 16

Private Type Iterator
    Type As DictionaryData
    StartIndex As Long
    EndIndex As Long
    Count As Long
    Index As Long
    HasLoaded As Boolean
    Current As IGeneric
    Keys As GenericArray
    Values As GenericArray
    DescriptorKeys As Long
    DescriptorValues As Long
End Type

Private Type Member
    Keys As GenericArray
    Values As GenericArray
    Size As Long
    IsReadOnly As Boolean
    SortOrder As SortOrder
    Iterator As Iterator
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get BaseIndex() As Long: BaseIndex = myclass.Keys.LowerBound: End Property
Public Property Get SortOrder() As SortOrder: SortOrder = myclass.SortOrder: End Property
Public Property Get Last() As IGeneric: Set Last = myclass.Values(myclass.Size): End Property
Public Property Get Count() As Long: Count = myclass.Size: End Property

Public Property Get GetKey(ByVal Index As Long) As IGeneric
    
    If (Index > myclass.Size) Then _
        Call Skynet.Throw(Me, "Public Property Get GetKey").IndexOutofRange("Index not valid")
    
    Set GetKey = myclass.Keys(Index)

End Property

Public Property Set ElementAt(ByVal Index As Long, ByVal Item As IGeneric)
    
    If (Index > myclass.Size) Then _
        Call Skynet.Throw(Me, "Public Property Get ElementAt").IndexOutofRange("Index not valid")
    
    Set myclass.Values(Index) = Item
    
End Property

Public Property Get ElementAt(ByVal Index As Long) As IGeneric

    If (Index > myclass.Size) Then _
        Call Skynet.Throw(Me, "Public Property Get ElementAt").IndexOutofRange("Index not valid")
    
    Set ElementAt = myclass.Values(Index)
    
End Property

Public Property Set Item(ByVal Key As IGeneric, ByVal Value As IGeneric)

    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Property set Item").KeyIsNothing
       
    Dim i As Long
    i = myclass.Keys.BinarySearch(Value:=Key, Index:=myclass.Keys.LowerBound, Length:=myclass.Size, Order:=myclass.SortOrder)
 
    If (i >= 0) Then _
        Set myclass.Values(i) = Value: Exit Property
    
    Call Insert(Index:=Not i, Key:=Key, Value:=Value)
    
 End Property
 
Public Property Get Item(ByVal Key As IGeneric) As IGeneric
Attribute Item.VB_UserMemId = 0

    Dim i As Long
    i = Me.IndexOfKey(Key)
    
    If (i >= 0) Then _
        Set Item = myclass.Values(i): Exit Function
        
    Call Skynet.Throw(Me, "Public Property Get Item").NoSuchElement
        
End Property

Public Property Get Capacity() As Long: Capacity = myclass.Keys.Length: End Property
Public Property Let Capacity(ByVal Value As Long)
    
    If (Value < myclass.Size) Then _
        Call Skynet.Throw(Me, "Public Property Let Capacity").ArgOutOfRange("Value < myClass.Size")
    
    If (myclass.Keys Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Property Let Capacity").NullPointer("myClass.Keys Is Nothing")
    
    If Not (Value = myclass.Keys.Length) Then
        If (Value > 0) Then
            Dim newKeys As GenericArray: Set newKeys = GenericArray.Build(Value)
            Dim newValues As GenericArray: Set newValues = GenericArray.Build(Value)
            
            If (myclass.Size > 0) Then
                With myclass.Keys
                    Call CopyMemory(dst:=ByVal newKeys.AdressOf(Element:=1), src:=ByVal .AdressOf(Element:=1), Length:=.TotalSize): _
                    Call ZeroMemory(dst:=ByVal .AdressOf(Element:=1), Length:=.TotalSize)
                End With
                With myclass.Values
                    Call CopyMemory(dst:=ByVal newValues.AdressOf(Element:=1), src:=ByVal .AdressOf(Element:=1), Length:=.TotalSize): _
                    Call ZeroMemory(dst:=ByVal .AdressOf(Element:=1), Length:=.TotalSize)
                End With
            End If
            
            Set myclass.Keys = newKeys
            Set myclass.Values = newValues
        Else
            Call myclass.Keys.Clear
            Call myclass.Values.Clear
        End If
    End If
End Property

Public Function Build(Optional ByVal Capacity As Long = DEFAULT_CAPACITY) As GenericSortedList
    
    If (Me Is GenericSortedList) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption
    
    If (Capacity < 0) Then _
        Call Skynet.Throw(Me, "Build").ArgOutOfRange("Capacity < 0")
    
    With myclass
        .SortOrder = Ascending
        Set .Keys = GenericArray.Build(Capacity)
        Set .Values = GenericArray.Build(Capacity)
    End With
   
    Set Build = Skynet.CreateInstance(New GenericSortedList, Virtual.Pointer, Virtual.Size)
    
End Function

Public Function BuildFrom(ByVal Dictionary As IGenericDictionary) As GenericSortedList
    
    If (Me Is GenericSortedList) = False Then _
        Call Skynet.Throw(Me, "Public Function BuildFrom").BuildExeption
    
    If (Dictionary Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function BuildFrom").NullPointer("(Dictionary Is Nothing)")
    
    With myclass
        .SortOrder = Ascending
          
        Set .Keys = GenericArray.Build(Dictionary.Count)
        Set .Values = GenericArray.Build(Dictionary.Count)
        
        Call Dictionary.CopyOf(t:=Keys_, GArray:=.Keys, ArrayIndex:=.Keys.LowerBound)
        Call Dictionary.CopyOf(t:=Values_, GArray:=.Values, ArrayIndex:=.Values.LowerBound)
        
        Call GenericArray.SortGenericArrays(Keys:=.Keys, Items:=.Values, Index:=.Keys.LowerBound, Length:=.Keys.Length, Order:=myclass.SortOrder)
        .Size = Dictionary.Count
    
    End With
   
    Set BuildFrom = Skynet.CreateInstance(New GenericSortedList, Virtual.Pointer, Virtual.Size)
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
End Sub

Private Sub Class_Terminate()
    Call Skynet.Dispose(Me)
End Sub

'Public Function GetKeys() As GenericArray: Set GetKeys = CopyOf(Keys_, GenericArray.Build(myClass.Size), GenericArray.LowerBound, "Public Property Get GetKeys"): End Function
'Public Function GetValues() As GenericArray: Set GetValues = CopyOf(Values_, GenericArray.Build(myClass.Size), GenericArray.LowerBound, "Public Property Get GetValues"): End Function

Private Sub EnsureCapacity(ByVal Min As Long)
    
    Dim newCapacity As Long
    
    With myclass.Keys
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
     
    If myclass.Size < Int(CDbl(myclass.Keys.Length) * 0.9) Then _
        Me.Capacity = myclass.Size

End Sub

Public Function TryGetValue(ByVal Key As IGeneric, ByRef Value As IGeneric) As Boolean

    Dim i As Long
    i = Me.IndexOfKey(Key)
    
    If (i >= 0) Then _
        Set Value = myclass.Values(i): TryGetValue = True
    
End Function

Public Function ContainsKey(ByVal Key As IGeneric) As Long: ContainsKey = Me.IndexOfKey(Key) > -1: End Function
Public Function ContainsValue(ByVal Value As IGeneric) As Long: ContainsValue = Me.IndexOfValue(Value) > -1: End Function

Public Function Remove(ByVal Key As IGeneric) As Boolean
    Dim i As Long
    i = Me.IndexOfKey(Key)
    
    If (i >= 0) Then _
        Call Me.RemoveAt(i): Remove = True
    
End Function

Public Sub RemoveAt(ByVal Index As Long)

    If (Index < myclass.Keys.LowerBound) Or (Index > myclass.Size) Then _
        Call Skynet.Throw(Me, "Public Sub RemoveAt").ArgOutOfRange("Index < myClass.Keys.LowerBound) Or (Index > myClass.Size)")
   
    Set myclass.Keys(Index) = Nothing
    Set myclass.Values(Index) = Nothing
    
    InterlockedDecrement myclass.Size
    
    If (Index < myclass.Size + myclass.Keys.LowerBound) Then
        With myclass.Keys
            Call CopyMemory(dst:=ByVal .AdressOf(Element:=Index), src:=ByVal .AdressOf(Element:=Index + 1), Length:=.SizeOf(myclass.Size - Index + .LowerBound)): _
            Call ZeroMemory(dst:=ByVal .AdressOf(Element:=myclass.Size), Length:=.SizeOf(1))
        End With
        With myclass.Values
            Call CopyMemory(dst:=ByVal .AdressOf(Element:=Index), src:=ByVal .AdressOf(Element:=Index + 1), Length:=.SizeOf(myclass.Size - Index + .LowerBound)): _
            Call ZeroMemory(dst:=ByVal .AdressOf(Element:=myclass.Size), Length:=.SizeOf(1))
        End With
    End If
    
End Sub

Public Sub Add(ByVal Key As IGeneric, ByVal Value As IGeneric)

    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Add").KeyIsNothing
       
    Dim i As Long
    i = myclass.Keys.BinarySearch(Value:=Key, Index:=1, Length:=myclass.Size, Order:=myclass.SortOrder)
    
    If (i >= 0) Then _
        Call Skynet.Throw(Me, "Public Sub Add").InvalidInput("Key duplicated")

    Call Insert(Index:=Not i, Key:=Key, Value:=Value)

End Sub

Private Sub Insert(ByVal Index As Long, ByVal Key As IGeneric, ByVal Value As IGeneric)

    If (myclass.Size = myclass.Keys.Length) Then _
        Call EnsureCapacity(myclass.Size + 1)

    If (Index < myclass.Size + myclass.Keys.LowerBound) Then
        With myclass.Keys
            Call CopyMemory(dst:=ByVal .AdressOf(Element:=Index + 1), src:=ByVal .AdressOf(Element:=Index), Length:=.SizeOf(myclass.Size - Index + .LowerBound)): _
            Call ZeroMemory(dst:=ByVal .AdressOf(Element:=Index), Length:=.SizeOf(1))
        End With
        With myclass.Values
            Call CopyMemory(dst:=ByVal .AdressOf(Element:=Index + 1), src:=ByVal .AdressOf(Element:=Index), Length:=.SizeOf(myclass.Size - Index + .LowerBound)): _
            Call ZeroMemory(dst:=ByVal .AdressOf(Element:=Index), Length:=.SizeOf(1))
        End With
    End If
    
    Set myclass.Keys(Index) = Key
    Set myclass.Values(Index) = Value
    Call InterlockedIncrement(myclass.Size)

End Sub

'Returns the index of the entry with a given key in this sorted list. The
'key is located through a binary search, and thus the average execution
'time of this method is proportional to Log2(Size), where
'Size is the Size of this sorted list. The returned value is -1 if
'the given key does not occur in this sorted list. Null is an invalid
'key value.
Public Function IndexOfKey(ByVal Key As IGeneric) As Long
     
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function IndexOfKey").KeyIsNothing
    
    Dim ret As Long
    ret = myclass.Keys.BinarySearch(Value:=Key, Index:=myclass.Keys.LowerBound, Length:=myclass.Size, Order:=myclass.SortOrder)

    IndexOfKey = IIf(ret >= 0, ret, -1)
    
End Function
'Returns the index of the first occurrence of an entry with a given value
'in this sorted list. The entry is located through a linear search, and
'thus the average execution time of this method is proportional to the
'Size of this sorted list. The elements of the list are compared to the
'given value using the IGeneric.Equals method.
Public Function IndexOfValue(ByVal Value As IGeneric) As Long
    IndexOfValue = myclass.Values.IndexOf(Value:=Value, StartIndex:=myclass.Values.LowerBound, Count:=myclass.Size)
End Function

Public Sub Clear():
    Call myclass.Keys.Clear
    Call myclass.Values.Clear
    myclass.Size = 0
End Sub

Private Function CopyOf(ByVal t As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long, ByVal Caller As String) As GenericArray
    
    If (GArray Is Nothing) Then _
        Call Skynet.Throw(Me, Caller).NullPointer("GArray Is Nothing")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call Skynet.Throw(Me, Caller).IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < myclass.Size) Then _
        Call Skynet.Throw(Me, Caller).InvalidInput("(GArray.Length - Index < myClass.Size)")
    
    Dim i As Long
     
    Select Case True
        Case t = Keys_
            Call GenericArray.Copy(SrcArray:=myclass.Keys, SrcIndex:=myclass.Keys.LowerBound, DstArray:=GArray, DstIndex:=ArrayIndex, Elements:=myclass.Size)
        Case t = Values_
            Call GenericArray.Copy(SrcArray:=myclass.Values, SrcIndex:=myclass.Values.LowerBound, DstArray:=GArray, DstIndex:=ArrayIndex, Elements:=myclass.Size)
        Case Else
            For i = myclass.Keys.LowerBound To myclass.Size
                Set GArray(ArrayIndex) = GenericPair(myclass.Keys(i), myclass.Values(i)): Call InterlockedIncrement(ArrayIndex)
            Next
    End Select
    
    Set CopyOf = GArray
    
End Function

Public Function Iterator(ByVal t As DictionaryData, Optional ByVal Index As Long, Optional ByVal Count As Long) As IGenericIterator
    
    Dim c As Member
    
    If Index = 0 Then _
        Index = myclass.Keys.LowerBound
        
    If Count = 0 Then _
        Count = myclass.Size
    
    If (Index < myclass.Keys.LowerBound) Or (Index > myclass.Size + myclass.Keys.LowerBound) Then _
        Call Skynet.Throw(Me, "Private Sub LoadIterator").IndexOutofRange("(Index < myClass.Keys.LowerBound) Or (Index > myClass.Size + myClass.Keys.LowerBound)")

    If (Count < 0) Or (Count > myclass.Size - Index + myclass.Keys.LowerBound) Then _
        Call Skynet.Throw(Me, "Public Function Iterator").ArgOutOfRange("(Count < 0) Or (Count > myClass.Size - Index + myClass.Keys.LowerBound)")
    
    Call LoadIterator(c.Iterator, t, myclass.Keys, myclass.Values, Index, Count)
    Set Iterator = Skynet.CreateInstance(New GenericSortedList, VarPtr(c), LenB(c))

'    Select Case True
'        Case t = Keys_
'            Set Iterator = myClass.Keys.Iterator(StartIndex:=myClass.Keys.LowerBound, Count:=myClass.Size)
'        Case t = Values_
'            Set Iterator = myClass.Values.Iterator(StartIndex:=myClass.Values.LowerBound, Count:=myClass.Size)
'        Case Else
'            Call LoadIterator(c.Iterator, Pairs_, myClass.Keys, myClass.Values, Index, Count)
'            Set Iterator = Skynet.CreateInstance(New GenericSortedList, VarPtr(c), LenB(c))
'    End Select
'
End Function

Private Sub LoadIterator( _
                            ByRef Iterator As Iterator, _
                            ByVal t As DictionaryData, _
                            ByVal Keys As GenericArray, _
                            ByVal Values As GenericArray, _
                            ByVal StartIndex As Long, _
                            ByVal Count As Long)
 
    With Iterator
        .Type = t
        Set .Current = Nothing
        Set .Keys = Keys
        Set .Values = Values
        .DescriptorKeys = Keys.Descriptor
        .DescriptorValues = Values.Descriptor
        .StartIndex = StartIndex
        .Count = Count
        .Index = .StartIndex - 1
        .EndIndex = .Index + .Count
        .HasLoaded = True
    End With
    
End Sub
'IGeneric
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GenericSortedList: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = Virtual.Pointer: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = Virtual.Size: End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericSortedList):  End Property

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member
    
    With c
        Set .Keys = Skynet.Clone(myclass.Keys)
        Set .Values = Skynet.Clone(myclass.Values)
        .Size = myclass.Size
        .SortOrder = myclass.SortOrder
    End With
    
    Set IGeneric_Clone = Skynet.CreateInstance(New GenericSortedList, VarPtr(c), LenB(c))
  
End Function

Private Sub IGeneric_Dispose()
    myclass = EmptyStructure
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    IGeneric_Equals = (Me Is Other)
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    Call Skynet.Throw(Me, "IObject_CompareTo").NotSupported
End Function

Private Function IGeneric_ToString() As String
    With myclass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Capacity: " & .Keys.Length & vbNewLine & _
                            "Elements: " & .Size & vbNewLine & _
                            "Sortorder: " & IIf(.SortOrder = Ascending, "Ascending", "Descending")
    End With
End Function

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_BaseIndex() As Long: IGenericReadOnlyList_BaseIndex = myclass.Keys.LowerBound: End Property
Private Property Get IGenericReadOnlyList_Count() As Long: IGenericReadOnlyList_Count = myclass.Size: End Property
Private Property Get IGenericReadOnlyList_ElementAt(ByVal Index As Long) As IGeneric
    
    If (Index > myclass.Size) Then _
        Call Skynet.Throw(Me, "Private Property Get IGenericReadOnlyList_ElementAt").IndexOutofRange("Index not valid")
    
    Set IGenericReadOnlyList_ElementAt = myclass.Values(Index)
    
End Property
Private Function IGenericReadOnlyList_IndexOf(ByVal Item As IGeneric) As Long: IGenericReadOnlyList_IndexOf = Me.IndexOfValue(Value:=Item): End Function
Private Sub IGenericReadOnlyList_Clear(): Call Me.Clear: End Sub
Private Sub IGenericReadOnlyList_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(Values_, GArray, ArrayIndex, "Private Sub IGenericReadOnlyList_CopyTo"): End Sub
Private Function IGenericReadOnlyList_Iterator() As IGenericIterator: Set IGenericReadOnlyList_Iterator = Me.Iterator(Values_): End Function

'IGenericDictionary
Private Property Set IGenericDictionary_Item(ByVal Key As IGeneric, ByVal Value As IGeneric): Set Me.Item(Key) = Value: End Property
Private Property Get IGenericDictionary_Item(ByVal Key As IGeneric) As IGeneric: Set IGenericDictionary_Item = Me.Item(Key): End Property
Private Property Get IGenericDictionary_Count() As Long: IGenericDictionary_Count = myclass.Size: End Property
Private Function IGenericDictionary_ContainsKey(ByVal Key As IGeneric) As Boolean: IGenericDictionary_ContainsKey = Me.IndexOfKey(Key) > -1: End Function
Private Function IGenericDictionary_ContainsValue(ByVal Value As IGeneric) As Boolean: IGenericDictionary_ContainsValue = Me.IndexOfValue(Value) > -1: End Function
Private Function IGenericDictionary_TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: IGenericDictionary_TryGetValue = Me.TryGetValue(Key, Value): End Function
Private Sub IGenericDictionary_Remove(ByVal Key As IGeneric): Call Me.Remove(Key): End Sub
Private Sub IGenericDictionary_Add(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value): End Sub
Private Sub IGenericDictionary_CopyOf(ByVal t As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(t:=t, GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="Private Sub IGenericDictionary_CopyOf"): End Sub
Private Sub IGenericDictionary_Clear(): Call Me.Clear: End Sub
Private Function IGenericDictionary_Iterator(ByVal t As DictionaryData) As IGenericIterator: Set IGenericDictionary_Iterator = Me.Iterator(t:=t): End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = myclass.IsReadOnly: End Property
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = CopyOf(Pairs_, GenericArray.Build(myclass.Size), GenericArray.LowerBound, "Private Function IGenericCollection_ToArray"): End Function
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.IndexOfKey(Key:=Item): End Function
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(Pairs_, GArray, ArrayIndex, "Private Sub IGenericCollection_CopyTo"): End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Iterator(Pairs_): End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.Current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    Dim K As IGeneric, V As IGeneric
     
    With myclass.Iterator
        If .Index = .EndIndex Then
            If .HasLoaded = False Then _
                Call Skynet.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator
            
            Set Out = Nothing
            Set .Current = Nothing
            Exit Function
        End If
        
        Call InterlockedIncrement(.Index)
        
        If .Type = Keys_ Then
            IGenericIterator_HasNext = (S_OK = SafeArrayGetElement(ByVal .DescriptorKeys, .Index, .Current))

        ElseIf .Type = Values_ Then
            IGenericIterator_HasNext = (S_OK = SafeArrayGetElement(ByVal .DescriptorValues, .Index, .Current))

        Else
            IGenericIterator_HasNext = S_OK = (SafeArrayGetElement(ByVal .DescriptorKeys, .Index, K) And SafeArrayGetElement(ByVal .DescriptorValues, .Index, V))
            Set .Current = GenericPair(K, V)
        End If
        
        If IGenericIterator_HasNext = True Then
            Set Out = .Current
        Else
            Set .Current = Nothing
            Set Out = Nothing
        End If
            
    End With
    
End Function

Private Sub IGenericIterator_Reset()
    With myclass.Iterator
        If .HasLoaded = False Then _
            Call Skynet.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
                
        Call LoadIterator(myclass.Iterator, .Type, .Keys, .Values, .StartIndex, .Count)
    End With
End Sub

