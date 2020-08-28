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
Implements IGenericEnumerator
Implements IPrototype

'OrderedHashTable offers IHashTable syntax with ordering.  Objects
'added or inserted in an IOrderedHashTable must have both a key and an index, and
'can be retrieved by either.
Private Const DEFAULT_CAPACITY As Long = 16

Private Type Member
    Internal As IGenericList
    HashTable As IGenericDictionary
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Function Build(Optional ByVal Capacity As Long = DEFAULT_CAPACITY) As GenericOrderedMap
    
    With myClass
        Set .Internal = GenericList.Build(Capacity)
        Set .HashTable = GenericMap.Build(Capacity)
    End With

    Set Build = Skynet.Create(GenericOrderedMap)

End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
End Sub
Public Property Get Entry(ByVal Index As Long) As GenericPair: Set Entry = myClass.Internal(Index): End Property
Public Property Get Count() As Long: Count = myClass.Internal.Elements.Count: End Property
Public Function GetKeys() As IGenericReadOnlyList

    Dim ga As GenericArray
    Set ga = GenericArray.Build(Me.Count)
    
    Dim i As Long
    For i = ga.LowerBound To ga.Length
        Set ga(i) = Me.Entry(i).Key
    Next
    
    Set GetKeys = ga.AsReadOnly
    
End Function

Public Function GetValues() As IGenericReadOnlyList
    
    Dim ga As GenericArray
    Set ga = GenericArray.Build(Me.Count)
    
    Dim i As Long
    For i = ga.LowerBound To ga.Length
        Set ga(i) = Me.Entry(i).Value
    Next
    
    Set GetValues = ga.AsReadOnly
    
End Function

Public Property Get Item(ByVal Key As IGeneric): Set Item = myClass.HashTable(Key): End Property
Attribute Item.VB_UserMemId = 0
Public Property Get ElementAt(ByVal Index As Long): Set ElementAt = Me.Entry(Index).Value: End Property

Public Property Set ElementAt(ByVal Index As Long, ByVal Item As IGeneric)
    Dim Key As IGeneric
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
        Call .Internal.Insert(.Internal.Elements.Count + 1, GenericPair(Key, Item))
    End With
End Sub

Public Function ContainsKey(ByVal Key As IGeneric) As Boolean: ContainsKey = myClass.HashTable.ContainsKey(Key): End Function
Public Function TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: TryGetValue = myClass.HashTable.TryGetValue(Key, Value): End Function
Private Function IndexOfKey(ByVal Key As IGeneric) As Long
    
    Dim i As Long
    With myClass
        For i = 1 To .Internal.Elements.Count
            If (Key.Equals(Me.Entry(i).Key)) Then _
                IndexOfKey = i: Exit Function
        Next
    End With
    
    IndexOfKey = -1

End Function

Public Function IndexOfValue(ByVal Value As IGeneric) As Long
    
    Dim i As Long
    With myClass
        For i = 1 To .Internal.Elements.Count
            If (Value.Equals(Me.Entry(i).Value)) Then _
                IndexOfValue = i: Exit Function
        Next
    End With
    
    IndexOfValue = -1

End Function

Public Sub Insert(ByVal Index As Long, ByVal Key As IGeneric, ByVal Value As IGeneric)
    
    If (Index > myClass.Internal.Elements.Count) Or (Index < 0) Then _
        Call Skynet.Throw(Me, "Public Sub Insert").IndexOutofRange("(Index > myClass.Internal.Elements.Count) Or (Index < 0)")
        
    With myClass
        Call .HashTable.Add(Key, Value)
        Call .Internal.Insert(Index, Key)
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
        
    Call Me.RemoveAt(IndexOfKey(Key))
End Sub

'IGeneric
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = IGeneric_HashValue = ObjPtr(Me): End Property
Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member
    Set c.HashTable = Skynet.Generic(myClass.HashTable).Clone
    Set c.Internal = Skynet.Generic(myClass.Internal).Clone
    
    Virtual.Pointer = VarPtr(c)
    Set IGeneric_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
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

Private Property Get IGenericEnumerator_Current() As IGeneric

End Property

Private Function IGenericEnumerator_MoveNext() As Boolean

End Function

Private Sub IGenericEnumerator_Reset()

End Sub

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_Elements() As IGenericCollection: Set IGenericReadOnlyList_Elements = Me: End Property
Private Function IGenericReadOnlyList_IndexOf(ByVal Item As IGeneric) As Long: IGenericReadOnlyList_IndexOf = Me.IndexOfValue(Item): End Function
Private Property Get IGenericReadOnlyList_Item(ByVal Index As Long) As IGeneric: Set IGenericReadOnlyList_Item = myClass.Internal(Index): End Property

'IGenericDictionary
Private Property Get IGenericDictionary_Elements() As IGenericCollection: Set IGenericDictionary_Elements = Me: End Property
Private Property Set IGenericDictionary_Item(ByVal Key As IGeneric, ByVal Item As IGeneric): Set Me.Item(Key) = Item: End Property
Private Property Get IGenericDictionary_Item(ByVal Key As IGeneric) As IGeneric: Set IGenericDictionary_Item = myClass.HashTable(Key): End Property
Private Property Get IGenericDictionary_Keys() As IGenericCollection: Set IGenericDictionary_Keys = Me.GetKeys: End Property
Private Property Get IGenericDictionary_Values() As IGenericCollection: Set IGenericDictionary_Values = Me.GetValues: End Property
Private Sub IGenericDictionary_Add(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value): End Sub
Private Sub IGenericDictionary_Remove(ByVal Key As IGeneric): Call Me.Remove(Key): End Sub
Private Function IGenericDictionary_TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean:  IGenericDictionary_TryGetValue = myClass.HashTable.TryGetValue(Key, Value): End Function
Private Function IGenericDictionary_ContainsKey(ByVal Key As IGeneric) As Boolean: IGenericDictionary_ContainsKey = myClass.HashTable.ContainsKey(Key): End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = False: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = False: End Property
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call myClass.Internal.Elements.CopyTo(GArray:=GArray, ArrayIndex:=ArrayIndex): End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = myClass.Internal.Elements.ToArray: End Function
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.IndexOfValue(Item) >= 0: End Function
Private Function IGenericCollection_NewEnumerator() As IGenericEnumerator

End Function

'IPrototype
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericOrderedMap: End Function
Private Property Get IPrototype_Default() As Object: Set IPrototype_Default = GenericOrderedMap: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property


