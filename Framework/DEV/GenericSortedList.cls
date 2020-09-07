VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSortedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A sorted list internally maintains two arrays that store the keys and
'values of the entries. The capacity of a sorted list is the allocated
'length of these internal arrays. As elements are added to a sorted list, the
'capacity of the sorted list is automatically increased as required by
'reallocating the internal arrays.  The capacity is never automatically
'decreased, but users can call either TrimExcess or
'Capacity explicitly.

'The GetKeys and GetValues methods of a sorted list
'provides access to the keys and values of the sorted list in the form of
'ReadonlyList implementations. The List objects returned by these
'methods are aliases for the underlying sorted list, so modifications
'made to those lists are directly reflected in the sorted list, and vice
'versa.

'The SortedList class provides a convenient way to create a sorted
'copy of another dictionary, such as a Hashtable. For example:

'Hashtable h = new Hashtable();
'h.Add(...);
'h.Add(...);

'SortedList s = new SortedList(h);

'The last line above creates a sorted list that contains a copy of the keys
'and values stored in the hashtable. In this particular example, the keys
'will be ordered according to the IComparable interface, which they
'all must implement. To impose a different ordering, SortedList also
'has a constructor that allows a specific IComparer implementation to
'be specified.

'@Folder("<T>Enumerable")
Option Compare Database

Option Explicit
'@PredeclaredId

Implements IPrototype
Implements IGeneric
Implements IGenericDictionary
Implements IGenericReadOnlyList
Implements IGenericCollection
Implements IGenericEnumerator

Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long

Private Const DEFAULT_CAPACITY As Long = 16

Private Enum InternalData
    Keys_ = 0
    Values_
    pairs_
End Enum

Private Type Enumerator
    Index As Long
    StartIndex As Long
    EndIndex As Long
    Current As IGeneric
    Parent As IGeneric
End Type

Private Type Member
    Keys As GenericArray
    Values As GenericArray
    Size As Long
    IsReadOnly As Boolean
    SortOrder As SortOrder
    Enumerator As Enumerator
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get SortOrder() As SortOrder: SortOrder = myClass.SortOrder: End Property
Public Property Get Last() As IGeneric: Set Last = myClass.Values(myClass.Size): End Property
Public Property Get GetKeys() As GenericArray: Set GetKeys = CopyOf(Keys_, GenericArray.Build(myClass.Size), GenericArray.LowerBound, "Public Property Get GetKeys"): End Property
Public Property Get GetValues() As GenericArray: Set GetValues = CopyOf(Values_, GenericArray.Build(myClass.Size), GenericArray.LowerBound, "Public Property Get GetValues"): End Property
Public Property Get Count() As Long: Count = myClass.Size: End Property

Public Property Set ElementAt(ByVal Index As Long, ByVal Item As IGeneric)
    
    If (Index > myClass.Size) Then _
        Call Skynet.Throw(Me, "Public Property Get ElementAt").IndexOutofRange("Index not valid")
    
    Set myClass.Values(Index) = Item
    
End Property

Public Property Get ElementAt(ByVal Index As Long) As IGeneric

    If (Index > myClass.Size) Then _
        Call Skynet.Throw(Me, "Public Property Get ElementAt").IndexOutofRange("Index not valid")
    
    Set ElementAt = myClass.Values(Index)
    
End Property

Public Property Set Item(ByVal Key As IGeneric, ByVal Value As IGeneric)

    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Property set Item").KeyIsNothing
       
    Dim i As Long
    i = myClass.Keys.BinarySearch(Value:=Key, Index:=myClass.Keys.LowerBound, Length:=myClass.Size, Order:=myClass.SortOrder)
 
    If (i >= 0) Then _
        Set myClass.Values(i) = Value: Exit Property
    
    Call Insert(Index:=Not i, Key:=Key, Value:=Value)
    
 End Property
 
Public Property Get Item(ByVal Key As IGeneric) As IGeneric
Attribute Item.VB_UserMemId = 0

    Dim i As Long
    i = Me.IndexOfKey(Key)
    
    If (i >= 0) Then _
        Set Item = myClass.Values(i): Exit Function
        
    Call Skynet.Throw(Me, "Public Property Get Item").NoSuchElement
        
End Property

Public Property Get Capacity() As Long: Capacity = myClass.Keys.Length: End Property
Public Property Let Capacity(ByVal Value As Long)
    
    If (Value < myClass.Size) Then _
        Call Skynet.Throw(Me, "Public Property Let Capacity").ArgOutOfRange("Value < myClass.Size")
    
    If (myClass.Keys Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Property Let Capacity").NullPointer("myClass.Keys Is Nothing")
    
    If Not (Value = myClass.Keys.Length) Then
        If (Value > 0) Then
            Dim newKeys As GenericArray: Set newKeys = GenericArray.Build(Value)
            Dim newValues As GenericArray: Set newValues = GenericArray.Build(Value)
            
            If (myClass.Size > 0) Then _
                Call Skynet.CopyData(Dst:=newKeys.AdressOf(Element:=1), Src:=myClass.Keys.AdressOf(Element:=1), Bytes:=myClass.Keys.TotalSize, ClearSource:=True): _
                Call Skynet.CopyData(Dst:=newValues.AdressOf(Element:=1), Src:=myClass.Values.AdressOf(Element:=1), Bytes:=myClass.Values.TotalSize, ClearSource:=True)

            Set myClass.Keys = newKeys
            Set myClass.Values = newValues

        Else
            Call myClass.Keys.Clear
            Call myClass.Values.Clear
        End If
    End If
End Property

Private Sub EnsureCapacity(ByVal Min As Long)
    
    Dim newCapacity As Long
    
    With myClass.Keys
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
     
    If myClass.Size < Int(CDbl(myClass.Keys.Length) * 0.9) Then _
        Me.Capacity = myClass.Size

End Sub

Public Function Build(Optional ByVal Capacity As Long = DEFAULT_CAPACITY) As GenericSortedList
    
    If (Capacity < 0) Then _
        Call Skynet.Throw(Me, "Build").ArgOutOfRange("Capacity < 0")
    
    With GenericArray
        Set myClass.Keys = .Build(Capacity)
        Set myClass.Values = .Build(Capacity)
    End With
    
    myClass.SortOrder = Skynet.SortOrder
    Set Build = Skynet.CreateInstance(New GenericSortedList, Virtual.Pointer, Virtual.Size)
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Virtual.Size = LenB(myClass)
End Sub

Private Sub Class_Terminate()
    Call Skynet.Dispose(Me)
End Sub

Public Function Enumerator(Optional ByVal StartIndex, Optional ByVal Count As Long) As IGenericEnumerator
    
    Dim C As Member
    Call LoadEnumerator(C.Enumerator, StartIndex, Count)
    Set C.Enumerator.Parent = Me 'make sure that parent cant go out of scope
    
    Set Enumerator = Skynet.CreateInstance(New GenericSortedList, VarPtr(C), LenB(C))

End Function

Public Function TryGetValue(ByVal Key As IGeneric, ByRef Value As IGeneric) As Boolean

    Dim i As Long
    i = Me.IndexOfKey(Key)
    
    If (i >= 0) Then _
        Set Value = myClass.Values(i): TryGetValue = True
    
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

    If (Index < myClass.Keys.LowerBound) Or (Index > myClass.Size) Then _
        Call Skynet.Throw(Me, "Public Sub RemoveAt").ArgOutOfRange("Index < myClass.Keys.LowerBound) Or (Index > myClass.Size)")
   
    Set myClass.Keys(Index) = Nothing
    Set myClass.Values(Index) = Nothing
    
    InterlockedDecrement myClass.Size
    
    If (Index < myClass.Size + myClass.Keys.LowerBound) Then
        With myClass.Keys
            Call Skynet.CopyData(Dst:=ByVal .AdressOf(Element:=Index), Src:=ByVal .AdressOf(Element:=Index + 1), Bytes:=.SizeOf(myClass.Size - Index + .LowerBound)): _
            Call Skynet.ClearData(Dst:=ByVal .AdressOf(Element:=myClass.Size), Bytes:=.SizeOf(1))
        End With
        With myClass.Values
            Call Skynet.CopyData(Dst:=ByVal .AdressOf(Element:=Index), Src:=ByVal .AdressOf(Element:=Index + 1), Bytes:=.SizeOf(myClass.Size - Index + .LowerBound)): _
            Call Skynet.ClearData(Dst:=ByVal .AdressOf(Element:=myClass.Size), Bytes:=.SizeOf(1))
        End With
    End If
    
End Sub

Public Sub Add(ByVal Key As IGeneric, ByVal Value As IGeneric)

    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Add").KeyIsNothing
       
    Dim i As Long
    i = myClass.Keys.BinarySearch(Value:=Key, Index:=1, Length:=myClass.Size, Order:=myClass.SortOrder)
    
    If (i >= 0) Then _
        Call Skynet.Throw(Me, "Public Sub Add").InvalidInput("Key duplicated")

    Call Insert(Index:=Not i, Key:=Key, Value:=Value)

End Sub

Private Sub Insert(ByVal Index As Long, ByVal Key As IGeneric, ByVal Value As IGeneric)

    If (myClass.Size = myClass.Keys.Length) Then _
        Call EnsureCapacity(myClass.Size + 1)

    If (Index < myClass.Size + myClass.Keys.LowerBound) Then
        With myClass.Keys
            Call Skynet.CopyData(Dst:=ByVal .AdressOf(Element:=Index + 1), Src:=ByVal .AdressOf(Element:=Index), Bytes:=.SizeOf(myClass.Size - Index + .LowerBound)): _
            Call Skynet.ClearData(Dst:=ByVal .AdressOf(Element:=Index), Bytes:=.SizeOf(1))
        End With
        With myClass.Values
            Call Skynet.CopyData(Dst:=ByVal .AdressOf(Element:=Index + 1), Src:=ByVal .AdressOf(Element:=Index), Bytes:=.SizeOf(myClass.Size - Index + .LowerBound)): _
            Call Skynet.ClearData(Dst:=ByVal .AdressOf(Element:=Index), Bytes:=.SizeOf(1))
        End With
    End If
    
    Set myClass.Keys(Index) = Key
    Set myClass.Values(Index) = Value
    Call InterlockedIncrement(myClass.Size)

End Sub

'Returns the index of the entry with a given key in this sorted list. The
'key is located through a binary search, and thus the average execution
'time of this method is proportional to Log2(size), where
'size is the size of this sorted list. The returned value is -1 if
'the given key does not occur in this sorted list. Null is an invalid
'key value.
Public Function IndexOfKey(ByVal Key As IGeneric) As Long
     
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function IndexOfKey").KeyIsNothing
    
    Dim ret As Long
    ret = myClass.Keys.BinarySearch(Value:=Key, Index:=myClass.Keys.LowerBound, Length:=myClass.Size, Order:=myClass.SortOrder)

    IndexOfKey = IIf(ret >= 0, ret, -1)
    
End Function
'Returns the index of the first occurrence of an entry with a given value
'in this sorted list. The entry is located through a linear search, and
'thus the average execution time of this method is proportional to the
'size of this sorted list. The elements of the list are compared to the
'given value using the IGeneric.Equals method.
Public Function IndexOfValue(ByVal Value As IGeneric) As Long
    IndexOfValue = myClass.Values.IndexOf(Value:=Value, StartIndex:=myClass.Values.LowerBound, Count:=myClass.Size)
End Function

Public Sub Clear():
    Call myClass.Keys.Clear
    Call myClass.Values.Clear
    myClass.Size = 0
End Sub

Private Sub LoadEnumerator(ByRef e As Enumerator, ByVal StartIndex As Long, ByVal EndIndex As Long)
    
    If (StartIndex < myClass.Keys.LowerBound) Or (StartIndex > myClass.Size + myClass.Keys.LowerBound) Then _
        Call Skynet.Throw(Me, "Private Sub LoadEnumerator").IndexOutofRange("(StartIndex < myClass.Keys.LowerBound) Or (StartIndex > myClass.Size + myClass.Keys.LowerBound)")

    If (Count < 0) Or (Count > myClass.Size - StartIndex + myClass.Keys.LowerBound) Then _
        Call Skynet.Throw(Me, "Private Sub LoadEnumerator").IndexOutofRange("(Count < 0) Or (Count > myClass.Size - StartIndex + myClass.Keys.LowerBound)")

    With e
        Set .Parent = e.Parent '! default is nothing, should only be set with newEnumerator
        .StartIndex = StartIndex
        .EndIndex = EndIndex
        .Index = StartIndex - 1
    End With

End Sub

Private Function CopyOf(ByVal Internal As InternalData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long, ByVal Caller As String) As GenericArray
    
    If (GArray Is Nothing) Then _
        Call Skynet.Throw(Me, Caller).NullPointer("GArray Is Nothing")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call Skynet.Throw(Me, Caller).IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < myClass.Size) Then _
        Call Skynet.Throw(Me, Caller).InvalidInput("(GArray.Length - Index < myClass.Size)")
    
    Dim i As Long
     
    Select Case True
        Case Internal = Keys_
            Call GenericArray.Copy(SrcArray:=myClass.Keys, SrcIndex:=myClass.Keys.LowerBound, DstArray:=GArray, DstIndex:=ArrayIndex, Elements:=myClass.Size)
        Case Internal = Values_
            Call GenericArray.Copy(SrcArray:=myClass.Values, SrcIndex:=myClass.Values.LowerBound, DstArray:=GArray, DstIndex:=ArrayIndex, Elements:=myClass.Size)
        Case Else
            For i = myClass.Keys.LowerBound To myClass.Size
                Set GArray(ArrayIndex) = GenericPairXY(myClass.Keys(i), myClass.Values(i)): Call InterlockedIncrement(ArrayIndex)
            Next
    End Select
    
    Set CopyOf = GArray
    
End Function

'IGeneric
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = IGeneric_HashValue = ObjPtr(Me): End Property
Private Function IGeneric_Clone() As IGeneric
    
    Dim C As Member
    Set C.Keys = Skynet.Clone(myClass.Keys)
    Set C.Values = Skynet.Clone(myClass.Values)
    C.Size = myClass.Size
    C.SortOrder = myClass.SortOrder
   
    Set IGeneric_Clone = Skynet.CreateInstance(New GenericSortedList, VarPtr(C), LenB(C))
  
End Function

Private Sub IGeneric_Dispose()
    myClass = EmptyStructure
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    IGeneric_Equals = (Me Is Other)
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    Call Skynet.Throw(Me, "IObject_CompareTo").NotSupported
End Function
Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericSortedList): End Function
Private Function IGeneric_ToString() As String
    With myClass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Capacity: " & .Keys.Length & vbNewLine & _
                            "Elements: " & .Size & vbNewLine & _
                            "Sortorder: " & IIf(.SortOrder = Ascending, "Ascending", "Descending")
    End With
End Function

'IGenericDictionary
Private Property Get IGenericDictionary_Elements() As IGenericCollection: Set IGenericDictionary_Elements = Me: End Property
Private Property Set IGenericDictionary_Item(ByVal Key As IGeneric, ByVal Value As IGeneric): Set Me.Item(Key) = Value: End Property
Private Property Get IGenericDictionary_Item(ByVal Key As IGeneric) As IGeneric: Set IGenericDictionary_Item = Me.Item(Key): End Property
Private Function IGenericDictionary_ContainsKey(ByVal Key As IGeneric) As Boolean: IGenericDictionary_ContainsKey = Me.IndexOfKey(Key) >= 0: End Function
Private Function IGenericDictionary_TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: IGenericDictionary_TryGetValue = Me.TryGetValue(Key, Value): End Function
Private Sub IGenericDictionary_Remove(ByVal Key As IGeneric): Call Me.Remove(Key): End Sub
Private Sub IGenericDictionary_Add(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value): End Sub
Private Sub IGenericDictionary_CopyKeysTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(Keys_, GArray, ArrayIndex, "Private Sub IGenericDictionary_CopyKeysTo"): End Sub
Private Sub IGenericDictionary_CopyValuesTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(Values_, GArray, ArrayIndex, "Private Sub IGenericDictionary_CopyValuesTo"): End Sub

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_Elements() As IGenericCollection: Set IGenericReadOnlyList_Elements = Me: End Property
Private Function IGenericReadOnlyList_IndexOf(ByVal Item As IGeneric) As Long: IGenericReadOnlyList_IndexOf = Me.IndexOfValue(Value:=Item): End Function
Private Property Get IGenericReadOnlyList_Item(ByVal Index As Long) As IGeneric
    
     If (Index > myClass.Size) Then _
        Call Skynet.Throw(Me, "Private Property Get IGenericReadOnlyList_Item").IndexOutofRange("Index not valid")
    
    Set IGenericReadOnlyList_Item = GenericPairXY(myClass.Keys(Index), myClass.Values(Index))
    
End Property

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = False: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = myClass.IsReadOnly: End Property
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = CopyOf(pairs_, GenericArray.Build(myClass.Size), GenericArray.LowerBound, "Private Function IGenericCollection_ToArray"): End Function
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.IndexOfKey(Key:=Item): End Function
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(pairs_, GArray, ArrayIndex, "Private Sub IGenericCollection_CopyTo"): End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_NewEnumerator() As IGenericEnumerator: Set IGenericCollection_NewEnumerator = Me.Enumerator: End Function
Private Function IGenericList_NewEnumerator() As IGenericEnumerator: Set IGenericList_NewEnumerator = Me.Enumerator: End Function

'IGenericEnumerator
Private Property Get IGenericEnumerator_Current() As IGeneric: Set IGenericEnumerator_Current = myClass.Enumerator.Current: End Property
Private Function IGenericEnumerator_MoveNext() As Boolean
    
    With myClass.Enumerator
        If .Index > .EndIndex Then _
            Set .Current = Nothing: Exit Function
        
        Call InterlockedIncrement(.Index)
        Set .Current = GenericPairXY(myClass.Keys(.Index), myClass.Values(.Index))
      
    End With
    
    IGenericEnumerator_MoveNext = True

End Function

Private Sub IGenericEnumerator_Reset()
    Call LoadEnumerator(myClass.Enumerator, myClass.Enumerator.StartIndex, myClass.Enumerator.EndIndex)
End Sub

'IPrototype
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericSortedList: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GenericSortedList: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.Size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property



