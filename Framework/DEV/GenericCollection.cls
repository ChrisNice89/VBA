VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
Option Explicit
'@Folder "<T>Enumerable"

Implements IGeneric
Implements IGenericReadOnlyList
Implements IGenericHashTable
Implements IGenericCollection
Implements IPrototype

'OrderedDictionary offers IDictionary syntax with ordering.  Objects
'added or inserted in an IOrderedDictionary must have both a key and an index, and
'can be retrieved by either.
Private Const DEFAULT_CAPACITY As Long = 32

Private Type Member
    KeyList As IGenericList
    Dictionary As IGenericHashTable
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Function Build(Optional ByVal Capacity As Long = DEFAULT_CAPACITY)
    
    With myClass
        Set .KeyList = GenericList.Build(Capacity)
        Set .Dictionary = GenericMap.Build(Capacity)
    End With

    Set Build = New GenericCollection

End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
End Sub

Public Property Get Count() As Long: Count = myClass.KeyList.Elements.Count: End Property
Public Property Get GetKeys() As IGenericCollection: Set GetKeys = GenericReadOnlyList.Build(myClass.KeyList): End Property
Public Property Get GetValues() As IGenericCollection: Set GetValues = GenericReadOnlyList.Build(Me): End Property

Public Property Get ElementAt(ByVal Index As Long)
    With myClass
        Set ElementAt = .Dictionary(.KeyList(Index))
    End With
End Property

Public Property Set ElementAt(ByVal Index As Long, ByVal Item As IGeneric)
    With myClass
        Set .Dictionary(.KeyList(Index)) = Item
    End With
End Property

Public Property Get ElementBy(ByVal Key As IGeneric)
    With myClass
        Set ElementBy = .Dictionary(Key)
    End With
End Property

Public Property Set ElementBy(ByVal Key As IGeneric, ByVal Item As IGeneric)
    With myClass
        If (.Dictionary.ContainsKey(Key)) Then _
            Set .KeyList(Me.IndexOfKey(Key)) = Key

        Set .Dictionary(Key) = Item
    End With
End Property

Public Sub Add(ByVal Key As IGeneric, ByVal Value As IGeneric)
    With myClass
        Call .Dictionary.Add(Key, Value)
        Call .KeyList.Insert(.KeyList.Elements.Count + 1, Key)
    End With
End Sub

Public Function TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: TryGetValue = myClass.Dictionary.TryGetValue(Key, Value): End Function
Public Function IndexOfKey(ByVal Key As IGeneric) As Long: IndexOfKey = myClass.KeyList.IndexOf(Key): End Function
Public Function IndexOfValue(ByVal Item As IGeneric) As Long
    
    Dim i As Long
    With myClass
        For i = 1 To .KeyList.Elements.Count
            If (Item.Equals(.Dictionary(.KeyList(i)))) Then _
                IndexOfValue = i: Exit Function
        Next
    End With
    
    IndexOfValue = -1

End Function

Public Function ContainsKey(ByVal Key As IGeneric) As Boolean: ContainsKey = myClass.Dictionary.ContainsKey(Key): End Function

Public Sub Insert(ByVal Index As Long, ByVal Key As IGeneric, ByVal Value As IGeneric)
    
    If (Index > myClass.KeyList.Elements.Count) Or (Index < 0) Then _
        Call Skynet.Push(Me, "Public Sub Insert").IndexOutofRange.Raise("(Index > myClass.KeyList.Elements.Count) Or (Index < 0)")
        
    With myClass
        Call .Dictionary.Add(Key, Value)
        Call .KeyList.Insert(Index, Key)
    End With
    
End Sub

Public Sub Clear()
    Call Skynet.Dispose(Me)
End Sub

Public Sub RemoveAt(ByVal Index As Long)
    
    If (Index < 0) Then _
        Exit Sub
        
    With myClass
        Call .Dictionary.Remove(.KeyList(Index))
        Call .KeyList.RemoveAt(Index)
    End With

End Sub

Public Sub Remove(ByVal Key As IGeneric)

    If (Key Is Nothing) Then _
        Call Skynet.Push(Me, "Public Sub Remove").NullPointer.Raise("Key is nothing")
        
    Call Me.RemoveAt(Me.IndexOfKey(Key))
End Sub

'IGeneric
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = IGeneric_HashValue = ObjPtr(Me): End Property
Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member
    Set c.Dictionary = Skynet.Generic(myClass.Dictionary).Clone
    Set c.KeyList = Skynet.Generic(myClass.KeyList).Clone
    
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
    Call Skynet.Push(Me, "IObject_CompareTo").NotSupported.Raise
End Function
Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericCollection): End Function
Private Function IGeneric_ToString() As String
    With myClass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Elements: " & Me.Count
    End With
End Function

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_Elements() As IGenericCollection: Set IGenericReadOnlyList_Elements = Me: End Property
Private Function IGenericReadOnlyList_IndexOf(ByVal Item As IGeneric) As Long: IGenericReadOnlyList_IndexOf = Me.IndexOfValue(Item): End Function
Private Property Get IGenericReadOnlyList_Item(ByVal Index As Long) As IGeneric: Set IGenericReadOnlyList_Item = Me.ElementAt(Index): End Property

'IGenericHashTable
Private Property Get IGenericHashTable_Elements() As IGenericCollection: Set IGenericHashTable_Elements = Me: End Property
Private Property Set IGenericHashTable_Item(ByVal Key As IGeneric, ByVal Item As IGeneric): Set Me.ElementBy(Key) = Item: End Property
Private Property Get IGenericHashTable_Item(ByVal Key As IGeneric) As IGeneric: Set IGenericHashTable_Item = myClass.Dictionary(Key): End Property
Private Property Get IGenericHashTable_Keys() As IGenericCollection: Set IGenericHashTable_Keys = Me.GetKeys: End Property
Private Property Get IGenericHashTable_Values() As IGenericCollection: Set IGenericHashTable_Values = myClass.Dictionary.Values: End Property
Private Sub IGenericHashTable_Add(ByVal Key As IGeneric, ByVal Value As IGeneric)
    With myClass
        Call .Dictionary.Add(Key, Value)
        Call .KeyList.Insert(.KeyList.Elements.Count, Key)
    End With
End Sub
Private Sub IGenericHashTable_Remove(ByVal Key As IGeneric): Call Me.Remove(Key): End Sub
Private Function IGenericHashTable_TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean:  IGenericHashTable_TryGetValue = myClass.Dictionary.TryGetValue(Key, Value): End Function
Private Function IGenericHashTable_ContainsKey(ByVal Key As IGeneric) As Boolean: IGenericHashTable_ContainsKey = myClass.Dictionary.ContainsKey(Key): End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = False: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = False: End Property
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call myClass.Dictionary.Elements.CopyTo(GArray:=GArray, ArrayIndex:=ArrayIndex): End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = myClass.Dictionary.Values: End Function
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.IndexOfValue(Item) >= 0: End Function
Private Function IGenericCollection_NewEnumerator() As IGenericEnumerator

End Function

'IPrototype
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericCollection: End Function
Private Property Get IPrototype_Default() As Object: Set IPrototype_Default = GenericCollection: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property


