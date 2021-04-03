VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericOrderedMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericOrderedMap
'*
'* Purpose: GenericOrderedMap represents a collection of key/value pairs that are accessible by the key or index and offers GenericMap syntax.
'*          In order to fulfill all of the operations, we require two copies of every TKey and TValue entry - one set in the list,
'*          and the other in the dictionary. This is required. There is no other way to efficiently provide all of the operations
'*          without mirroring all of the data between both containers. The trade off is more speed at the cost of more memory.
'*
'*          It is implemented by storing elements in two separate internal structures: a Hashtable,
'*          which stores the key/value pairs like a normal Map and an GenericList which stores pair structures containing the key/value pairs.
'*
'*          This implementation of an ordered map is very good at lookup operations:
'*          the list allows O(1) lookups by index, and the map allows O(1) lookups by key.
'*          However, the necessity of keeping the list synchronized with the map means that insert/delete operations have the performance disadvantage
'*          of performing those operations on an array (O(n) at worst).
'*          There is also, of course, the extra memory requirement of storing both data structures.
'*          Because of these disadvantages, GenericOrderedMap should only be used when insert/delete operations will be minimal
'*          and there is a need to efficiently access elements by index and/or key.
'*
'*          Generally, manipulating the GenericOrderedMap by index is faster than manipulating it by key;
'*          even if the algorithmic complexity is the same, the actual running time will be less when using an index.
'*          This is because every modification to an OrderedDictionary requires accessing both the list and the map and converting an index to its
'*          corresponding key is a O(1) operation (it's just a lookup in the list),
'*          but converting a key to its corresponding index is a O(n) operation, since it requires a linear search through the list to find the key.
'*          Put another way, the list is aware of the dictionary's identifier (the key), but the map is not aware of the list's identifier (the index).
'*          Therefore, you should try to manipulate the GenericOrderedMap  by index where possible.
'*
'* Notes:   The elements of an GenericOrderedMap are not sorted by the key.
'*          Elements can be accessed either by the key or by the index.
'*          Each element is a key/value pair stored in a GenericPair object.
'*          Offers IGenericMap syntax
'*
'*======================================================================================================================

Implements IGeneric
Implements IGenericMap
Implements IGenericReadOnlyList
Implements IGenericCollection

Private Const DEFAULT_CAPACITY As Long = 16

Private Type Member
    Comparer As IGenericComparer
    KeyList As GenericOrderedList
    LookupTable As IGenericMap
End Type
Private myclass As Member

Public Property Get Elements() As IGenericCollection: Set Elements = Me: End Property
Public Property Get Map() As IGenericMap: Set Map = Me: End Property
Public Property Get Comparer() As IGenericComparer: Set Comparer = myclass.Comparer: End Property
Public Property Get First() As Long: First = myclass.KeyList.First: End Property
Public Property Get Last() As Long: Last = myclass.KeyList.Last: End Property

Public Property Get GetKey(ByVal Index As Long) As IGeneric
    
    If (Index >= myclass.KeyList.Elements.Count) Then _
        Call System.Throw(Me, "Public Property Get GetKey").IndexOutofRange("Index not valid")
    
    Set GetKey = myclass.KeyList.ElementAt(Index)

End Property

Public Property Get Item(ByVal Key As IGeneric) As IGeneric: Set Item = myclass.LookupTable.Item(Key): End Property
Public Property Set Item(ByVal Key As IGeneric, ByVal Item As IGeneric)
    
    If (myclass.LookupTable.TryAdd(Key, Item, Overwrite:=True)) Then _
        Call myclass.KeyList.Add(Key)
        
End Property
Public Property Get ElementAt(ByVal Index As Long) As IGeneric: Set ElementAt = myclass.LookupTable.Item(Me.GetKey(Index)): End Property
Public Property Set ElementAt(ByVal Index As Long, ByVal Item As IGeneric): Set myclass.LookupTable.Item(Me.GetKey(Index)) = Item: End Property

Public Function Build(Optional ByVal Capacity As Long, Optional ByVal Comparer As IGenericComparer) As GenericOrderedMap
    
    Dim newMap As Member
    
    If (Capacity < 0) Then _
       Call System.Throw(Me, "Build").ArgOutOfRange("Capacity < 0")
    
    If Comparer Is Nothing Then _
        Set Comparer = IGenericComparer
    Set newMap.Comparer = Comparer
    
    Set newMap.KeyList = GenericOrderedList.Build(Capacity)
    Set newMap.LookupTable = GenericLinkedMap.Build(Capacity, newMap.Comparer)
    
    Set Build = System.CreateInstance(New GenericOrderedMap, VarPtr(newMap), LenB(newMap))

End Function

Public Function GetEntry(ByVal Index As Long) As GenericPair

    Dim Key As IGeneric
    Set Key = myclass.KeyList.ElementAt(Index)
    Set GetEntry = GenericPair.Build(Key, myclass.LookupTable.Item(Key))
    
End Function

Public Function GetKeys() As GenericArray: Set GetKeys = myclass.KeyList.Elements.ToArray: End Function
Public Function GetValues() As GenericArray: Set GetValues = myclass.LookupTable.GetValues: End Function

Public Sub Add(ByVal Key As IGeneric, ByVal Value As IGeneric, Optional ByVal Overwrite As Boolean)
    
    If (myclass.LookupTable.TryAdd(Key, Value, Overwrite:=Overwrite)) Then _
        Call myclass.KeyList.Add(Key)
   
End Sub

Public Function AddAll(ByVal Map As GenericLinkedMap) As GenericOrderedMap
    
    Dim P As GenericPair
  
    With Map.Elements.Iterator()
        Do While .HasNext(P): Call Me.Add(P.Key, P.Value): Loop
    End With
    
    Set AddAll = Me
    
End Function

Public Function ContainsKey(ByVal Key As IGeneric) As Boolean: ContainsKey = myclass.LookupTable.ContainsKey(Key): End Function
Public Function ContainsValue(ByVal Value As IGeneric) As Boolean: ContainsValue = myclass.LookupTable.ContainsValue(Value): End Function

Public Function TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: TryGetValue = myclass.LookupTable.TryGetValue(Key, Value): End Function
Public Function IndexOfKey(ByVal Key As IGeneric) As Long: IndexOfKey = myclass.KeyList.IndexOf(Element:=Key, Comparer:=myclass.Comparer): End Function
'Use equality
Public Function IndexOfValue(ByVal Value As IGeneric) As Long
    
    Dim i As Long
    
    With myclass.KeyList
        For i = .First To .Last
            If (myclass.LookupTable.Item(.ElementAt(i)).Equals(Value)) Then _
                IndexOfValue = i: Exit Function
        Next
    End With
    IndexOfValue = -1

End Function

Public Sub RemoveAt(ByVal Index As Long)

    If (myclass.LookupTable.TryRemove(myclass.KeyList.ElementAt(Index))) Then _
        Call myclass.KeyList.RemoveAt(Index)
        
End Sub

Public Function TryRemove(ByVal Key As IGeneric) As Boolean
    
    Dim i As Long
    
    If (myclass.LookupTable.TryRemove(Key)) Then
        i = myclass.KeyList.IndexOf(Element:=Key, Comparer:=myclass.Comparer)
        Call myclass.KeyList.RemoveAt(Index:=i)
        TryRemove = True
    End If
    
End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericOrderedMap: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GenericOrderedMap: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericOrderedMap):  End Property
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
                            "Elements: " & Me.Elements.Count
    End With
End Function

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_Elements() As IGenericCollection: Set IGenericReadOnlyList_Elements = Me: End Property
Private Property Get IGenericReadOnlyList_First() As Long: IGenericReadOnlyList_First = myclass.KeyList.First: End Property
Private Property Get IGenericReadOnlyList_Last() As Long: IGenericReadOnlyList_Last = myclass.KeyList.Last: End Property

Private Function IGenericReadOnlyList_GetAt(ByVal Index As Long) As IGeneric
    
    If (Index >= myclass.KeyList.Elements.Count) Then _
        Call System.Throw(Me, "Private Function IGenericReadOnlyList_GetAt").IndexOutofRange("(Index >= myclass.KeyList.Elements.Count)")
    
    Set IGenericReadOnlyList_GetAt = Me.GetEntry(Index)
    
End Function

Private Function IGenericReadOnlyList_IndexOf(ByVal Element As IGeneric) As Long
  
    If myclass.LookupTable.ContainsKey(Key:=Element) Then
        IGenericReadOnlyList_IndexOf = myclass.KeyList.IndexOf(Element:=Element, Comparer:=myclass.Comparer)
    Else
        IGenericReadOnlyList_IndexOf = -1
    End If
    
End Function

'IGenericMap
Private Property Get IGenericMap_Elements() As IGenericCollection: Set IGenericMap_Elements = Me: End Property
Private Property Set IGenericMap_Item(ByVal Key As IGeneric, ByVal Item As IGeneric): Set Me.Item(Key) = Item: End Property
Private Property Get IGenericMap_Item(ByVal Key As IGeneric) As IGeneric: Set IGenericMap_Item = Me.Item(Key): End Property
Private Property Get IGenericMap_Comparer() As IGenericComparer: Set IGenericMap_Comparer = myclass.Comparer: End Property
'Todo
Private Function IGenericMap_TryAdd(ByVal Key As IGeneric, ByVal Value As IGeneric, Optional ByVal Overwrite As Boolean) As Boolean
    
    If (myclass.LookupTable.TryAdd(Key, Value, Overwrite:=Overwrite)) Then
        Call myclass.KeyList.Add(Key)
        IGenericMap_TryAdd = True
    End If
    
End Function

Private Function IGenericMap_TryRemove(ByVal Key As IGeneric) As Boolean: IGenericMap_TryRemove = Me.TryRemove(Key:=Key): End Function
Private Function IGenericMap_TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: IGenericMap_TryGetValue = myclass.LookupTable.TryGetValue(Key, Value): End Function

Private Function IGenericMap_GetKeys() As GenericArray: Set IGenericMap_GetKeys = Me.GetKeys: End Function
Private Function IGenericMap_GetValues() As GenericArray: Set IGenericMap_GetValues = Me.GetValues: End Function
Private Function IGenericMap_ContainsKey(ByVal Key As IGeneric) As Boolean: IGenericMap_ContainsKey = myclass.LookupTable.ContainsKey(Key): End Function
Private Function IGenericMap_ContainsValue(ByVal Value As IGeneric) As Boolean: IGenericMap_ContainsValue = myclass.LookupTable.ContainsValue(Value): End Function

Private Function IGenericMap_RemoveAll(ByVal Keys As IGenericCollection) As Long
    
    Dim Key As IGeneric
    Dim Count As Long
    
    With Keys.Iterator
        Do While .HasNext(Key)
            If Me.TryRemove(Key:=Key) Then _
                Count = Count + 1
            
        Loop
    End With
    
    IGenericMap_RemoveAll = Count
    
End Function

Private Function IGenericMap_AddAll(ByVal Map As IGenericMap, Optional ByVal Overwrite As Boolean) As IGenericMap
    
    Dim Pair As GenericPair
  
    If (Map Is Nothing) Then _
        Call System.Throw(Me, "Private Function IGenericMap_AddAll").IsNothing("(Map Is Nothing)")
    
    If (Map Is Me) Then _
        Exit Function
            
    If (Map.Elements.Count > 0) Then
        With Map.Elements.Iterator
            Do While .HasNext(Pair): Call Me.Add(Pair.Key, Pair.Value, Overwrite:=Overwrite): Loop
        End With
    End If
    
    Set IGenericMap_AddAll = Me
    
End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long

     If Not myclass.KeyList Is Nothing Then _
        IGenericCollection_Count = myclass.KeyList.Elements.Count

End Property
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long)

    If (GArray Is Nothing) Then _
        Call System.Throw(Me, "Public Sub CopyTo").IsNothing("(GArray Is Nothing)")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex >= GArray.Length) Then _
        Call System.Throw(Me, "Public Sub CopyTo").IndexOutofRange("(ArrayIndex < GArray.LowerBound) Or (ArrayIndex >= GArray.Length)")
        
    Call myclass.LookupTable.Elements.CopyTo(GArray:=GArray, ArrayIndex:=ArrayIndex)
    
End Sub

Private Sub IGenericCollection_Clear()
        
    If Not (myclass.KeyList Is Nothing) Then _
        Call myclass.KeyList.Elements.Clear
       
    If Not (myclass.LookupTable Is Nothing) Then _
        Call myclass.LookupTable.Elements.Clear

End Sub
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = myclass.LookupTable.Elements.ToArray: End Function
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = myclass.LookupTable.Elements.Contains(Element:=Item): End Function
Private Function IGenericCollection_ContainsAll(ByVal Collection As IGenericCollection) As Boolean: IGenericCollection_ContainsAll = myclass.LookupTable.Elements.ContainsAll(Collection:=Collection):: End Function
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = myclass.LookupTable.Elements.Iterator: End Function
Private Function IGenericCollection_Copy() As IGenericCollection
    
    Dim c As Member
    
    Set c.Comparer = myclass.Comparer
    Set c.LookupTable = myclass.LookupTable.Elements.Copy
    Set c.KeyList = myclass.KeyList.Elements.Copy
    
    Set IGenericCollection_Copy = System.CreateInstance(New GenericOrderedMap, VarPtr(c), LenB(c))
    
End Function
