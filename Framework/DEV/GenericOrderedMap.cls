VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericOrderedMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
Option Explicit
'@Folder "<T>Enumerable"

Implements IGeneric
Implements IGenericReadOnlyList
Implements IGenericDictionary
Implements IGenericCollection
Implements IGenericIterator
Implements IPrototype

Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long

'OrderedHashTable offers IHashTable syntax with ordering.  Objects
'added or inserted in an IOrderedHashTable must have both a key and an index, and
'can be retrieved by either.
Private Const DEFAULT_CAPACITY As Long = 16

Private Type Iterator
    Type As DictionaryData
    StartIndex As Long
    EndIndex As Long
    Count As Long
    Index As Long
    HasLoaded As Boolean
    Current As IGeneric
    Entries As GenericList
End Type

Private Type Member
    Internal As GenericList
    HashTable As IGenericDictionary
    Iterator As Iterator
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get BaseIndex() As Long: BaseIndex = myClass.Internal.BaseIndex: End Property
Public Property Get Entry(ByVal Index As Long) As GenericPair: Set Entry = myClass.Internal(Index): End Property
Public Property Get Count() As Long: Count = myClass.Internal.Count: End Property

Public Function Build(Optional ByVal Capacity As Long = DEFAULT_CAPACITY) As GenericOrderedMap
    
    If (Me Is GenericOrderedMap) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption
    
    With myClass
        Set .Internal = GenericList.Build(Capacity)
        Set .HashTable = GenericMap.Build(Capacity * 2) 'GenericTree.Build
    End With

    Set Build = Skynet.CreateInstance(New GenericOrderedMap, Virtual.Pointer, Virtual.Size)

End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Virtual.Size = LenB(myClass)
End Sub

Public Function GetKeys() As GenericArray

    ReDim Keys(GenericArray.LowerBound To myClass.Internal.Count) As IGeneric

    Dim i As Long

    For i = LBound(Keys) To UBound(Keys)
        Set Keys(i) = Me.Entry(i).Key
    Next
    
    Set GetKeys = GenericArray.BuildFrom(Keys)
    
End Function

Public Function GetValues() As GenericArray

    ReDim Values(GenericArray.LowerBound To myClass.Internal.Count) As IGeneric

    Dim i As Long
    
    For i = LBound(Values) To UBound(Values)
        Set Values(i) = Me.Entry(i).Value
    Next
    
    Set GetValues = GenericArray.BuildFrom(Values)

End Function

Public Property Get Item(ByVal Key As IGeneric): Set Item = myClass.HashTable(Key): End Property
Attribute Item.VB_UserMemId = 0
Public Property Get ElementAt(ByVal Index As Long): Set ElementAt = Me.Entry(Index).Value: End Property

Public Property Set ElementAt(ByVal Index As Long, ByVal Item As IGeneric)
    Dim Key As IGeneric
    
    If (Index > myClass.Internal.Count) Then _
        Call Skynet.Throw(Me, "Public Property Set ElementAt").IndexOutofRange("Index not valid")
    
    Set Key = Me.Entry(Index).Key
    Set myClass.HashTable(Key) = Item
    Set myClass.Internal(Index) = GenericPair(Key, Item)
    
End Property

Public Property Set Item(ByVal Key As IGeneric, ByVal Item As IGeneric)
    With myClass
        If (.HashTable.ContainsKey(Key)) Then _
            Set .Internal(IndexOfKey(Key)) = GenericPair(Key, Item)

        Set .HashTable(Key) = Item
    End With
End Property

Public Sub Add(ByVal Key As IGeneric, ByVal Item As IGeneric)
    With myClass
        Call .HashTable.Add(Key, Item)
        Call .Internal.Add(GenericPair(Key, Item))
    End With
End Sub

Public Sub AddAll(ByVal Map As IGenericDictionary)
    
    Dim p As GenericPair
  
    With Map.Iterator(T:=Pairs_) ' Copy
        Do While .HasNext(p): Call Me.Add(p.Key, p.Value): Loop
        'Do While .HasNext(p): Call Me.AddPair(p): Loop
    End With
    
End Sub

Public Sub AddPair(ByVal Pair As GenericPair)
    
    With myClass
        Call .HashTable.Add(Pair.Key, Pair.Value)
        Call .Internal.Add(Skynet.Clone(Pair))
    End With
    
End Sub
Public Function ContainsKey(ByVal Key As IGeneric) As Boolean: ContainsKey = myClass.HashTable.ContainsKey(Key): End Function
Public Function TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: TryGetValue = myClass.HashTable.TryGetValue(Key, Value): End Function
Public Function IndexOfKey(ByVal Key As IGeneric) As Long
    
    Dim i As Long
    With myClass
        For i = .Internal.BaseIndex To .Internal.Count
            If (Key.Equals(Me.Entry(i).Key)) Then _
                IndexOfKey = i: Exit Function
        Next
    End With
    IndexOfKey = -1

End Function

Public Function IndexOfValue(ByVal Value As IGeneric) As Long
    
    Dim i As Long
    With myClass
        For i = .Internal.BaseIndex To .Internal.Count
            If (Value.Equals(Me.Entry(i).Value)) Then _
                IndexOfValue = i: Exit Function
        Next
    End With
    
    IndexOfValue = -1

End Function

Public Sub Insert(ByVal Index As Long, ByVal Key As IGeneric, ByVal Value As IGeneric)
    
    If (Index > myClass.Internal.Count) Or (Index < 0) Then _
        Call Skynet.Throw(Me, "Public Sub Insert").IndexOutofRange("(Index > myClass.Internal.Count) Or (Index < 0)")
        
    With myClass
        Call .HashTable.Add(Key, Value)
        Call .Internal.Insert(Index, GenericPair(Key, Value))
    End With
    
End Sub

Public Sub Clear()
    Call Skynet.Dispose(Me)
End Sub

Public Sub RemoveAt(ByVal Index As Long)
    
    If (Index < 0) Then _
        Exit Sub
        
    With myClass
        Call .HashTable.Remove(Me.Entry(Index).Key)
        Call .Internal.RemoveAt(Index)
    End With

End Sub

Public Sub Remove(ByVal Key As IGeneric)

    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Sub Remove").KeyIsNothing
        
    Call Me.RemoveAt(Me.IndexOfKey(Key))
End Sub

Private Function CopyOf(ByVal T As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long, ByVal Caller As String) As GenericArray
    
    If (GArray Is Nothing) Then _
        Call Skynet.Throw(Me, Caller).NullPointer("GArray Is Nothing")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call Skynet.Throw(Me, Caller).IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < myClass.Internal.Count) Then _
        Call Skynet.Throw(Me, Caller).InvalidInput("(GArray.Length - ArrayIndex + GArray.LowerBound < myClass.Internal.Count)")
    
    Dim i As Long
    
    Select Case True
        Case T = keys_
            For i = 1 To myClass.Internal.Count
                Set GArray(ArrayIndex) = Me.Entry(i).Key: Call InterlockedIncrement(ArrayIndex)
            Next
        Case T = Values_
            For i = 1 To myClass.Internal.Count
                Set GArray(ArrayIndex) = Me.Entry(i).Value: Call InterlockedIncrement(ArrayIndex)
            Next
        Case Else
            Call myClass.Internal.CopyTo(Index:=GenericArray.LowerBound, GArray:=GArray, ArrayIndex:=ArrayIndex, elements:=myClass.Internal.Count)
    End Select
    
    Set CopyOf = GArray
    
End Function

Public Function Iterator(ByVal T As DictionaryData, Optional ByVal Index As Long, Optional ByVal Count As Long) As IGenericIterator
    
    Dim C As Member
    
    If Index = 0 Then _
        Index = myClass.Internal.BaseIndex
        
    If Count = 0 Then _
        Count = myClass.Internal.Count
    
    If (Index < myClass.Internal.BaseIndex) Or (Index > myClass.Internal.Count + myClass.Internal.BaseIndex) Then _
        Call Skynet.Throw(Me, "Private Sub LoadIterator").IndexOutofRange("(Index < myClass.Internal.BaseIndex) Or (Index > myClass.Internal.Count + myClass.Internal.BaseIndex)")

    If (Count < 0) Or (Count > myClass.Internal.Count - Index + myClass.Internal.BaseIndex) Then _
        Call Skynet.Throw(Me, "Public Function Iterator").ArgOutOfRange("(Count < 0) Or (Count > myClass.Internal.Count - Index + myClass.Internal.BaseIndex)")
    
    If T = Pairs_ Then
        Set Iterator = myClass.Internal.Iterator(Index:=Index, Count:=Count)
    Else
        Call LoadIterator(C.Iterator, T, myClass.Internal, Index, Count)
        Set Iterator = Skynet.CreateInstance(New GenericOrderedMap, VarPtr(C), LenB(C))
    End If
    
End Function

Private Sub LoadIterator( _
                            ByRef Iterator As Iterator, _
                            ByVal T As DictionaryData, _
                            ByVal Entries As GenericList, _
                            ByVal StartIndex As Long, _
                            ByVal Count As Long)
 
    With Iterator
        .Type = T
        Set .Current = Nothing
        Set .Entries = Entries
        .StartIndex = StartIndex
        .Count = Count
        .Index = .StartIndex - 1
        .EndIndex = .Index + .Count
        .HasLoaded = True
    End With
    
End Sub
'IGeneric
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = IGeneric_HashValue = ObjPtr(Me): End Property
Private Function IGeneric_Clone() As IGeneric
    
    Dim C As Member
    Set C.HashTable = Skynet.Clone(myClass.HashTable)
    Set C.Internal = Skynet.Clone(myClass.Internal)
    
    Set IGeneric_Clone = Skynet.CreateInstance(New GenericOrderedMap, VarPtr(C), Virtual.Size)
   
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
Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericOrderedMap): End Function
Private Function IGeneric_ToString() As String
    With myClass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Elements: " & Me.Count
    End With
End Function

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_BaseIndex() As Long: IGenericReadOnlyList_BaseIndex = myClass.Internal.BaseIndex: End Property
Private Property Get IGenericReadOnlyList_Count() As Long: IGenericReadOnlyList_Count = myClass.Internal.Count: End Property
Private Property Get IGenericReadOnlyList_Item(ByVal Index As Long) As IGeneric
    
    Dim p As GenericPair
    
    If (Index > myClass.Internal.Count) Then _
        Call Skynet.Throw(Me, "Private Property Get IGenericReadOnlyList_Item").IndexOutofRange("Index not valid")
    
    Set p = myClass.Internal(Index)
    
    Set IGenericReadOnlyList_Item = p.Value

End Property

Private Sub IGenericReadOnlyList_Clear(): Call Me.Clear: End Sub
Private Sub IGenericReadOnlyList_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(Values_, GArray, ArrayIndex, "Private Sub IGenericDictionary_CopyOf"): End Sub
Private Function IGenericReadOnlyList_Iterator() As IGenericIterator: Set IGenericReadOnlyList_Iterator = Me.Iterator(Values_): End Function
Private Function IGenericReadOnlyList_IndexOf(ByVal Item As IGeneric) As Long: IGenericReadOnlyList_IndexOf = Me.IndexOfValue(Item): End Function

'IGenericDictionary
Private Property Set IGenericDictionary_Item(ByVal Key As IGeneric, ByVal Item As IGeneric): Set Me.Item(Key) = Item: End Property
Private Property Get IGenericDictionary_Item(ByVal Key As IGeneric) As IGeneric: Set IGenericDictionary_Item = myClass.HashTable(Key): End Property
Private Property Get IGenericDictionary_Count() As Long: IGenericDictionary_Count = myClass.HashTable.Count: End Property
Private Function IGenericDictionary_TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean:  IGenericDictionary_TryGetValue = myClass.HashTable.TryGetValue(Key, Value): End Function
Private Function IGenericDictionary_ContainsKey(ByVal Key As IGeneric) As Boolean: IGenericDictionary_ContainsKey = myClass.HashTable.ContainsKey(Key): End Function
Private Function IGenericDictionary_ContainsValue(ByVal Value As IGeneric) As Boolean: IGenericDictionary_ContainsValue = Me.IndexOfValue(Value) > -1: End Function
Private Sub IGenericDictionary_Add(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value): End Sub
Private Sub IGenericDictionary_Remove(ByVal Key As IGeneric): Call Me.Remove(Key): End Sub
Private Sub IGenericDictionary_CopyOf(ByVal T As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(T:=T, GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="Private Sub IGenericDictionary_CopyOf"): End Sub
Private Function IGenericDictionary_Iterator(ByVal T As DictionaryData) As IGenericIterator: Set IGenericDictionary_Iterator = Me.Iterator(T:=T): End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = False: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = False: End Property
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(T:=Pairs_, GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="Private Sub IGenericCollection_CopyTo"): End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = myClass.Internal.ToArray: End Function
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = myClass.HashTable.ContainsKey(Item): End Function
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Iterator(T:=Pairs_): End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myClass.Iterator.Current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    Dim p As GenericPair
    
    With myClass.Iterator
        If .Index = .EndIndex Then
            If .HasLoaded = False Then _
                Call Skynet.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator
            
            Set Out = Nothing
            Set .Current = Nothing
            Exit Function
        End If
        
        IGenericIterator_HasNext = True
        
        Set p = .Entries(InterlockedIncrement(.Index))
        
        If .Type = keys_ Then _
            Set .Current = p.Key: Set Out = .Current: Exit Function
        
        If .Type = Values_ Then _
            Set .Current = p.Value: Set Out = .Current: Exit Function
        
        Set .Current = p: Set Out = .Current
        
    End With

End Function

Private Sub IGenericIterator_Reset()
    With myClass.Iterator
        If .HasLoaded = False Then _
            Call Skynet.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
                
        Call LoadIterator(myClass.Iterator, .Type, .Entries, .StartIndex, .Count)
    End With
End Sub

'IPrototype
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericOrderedMap: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GenericOrderedMap: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.Size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property
