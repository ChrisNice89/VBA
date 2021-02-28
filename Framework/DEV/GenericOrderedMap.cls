VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericOrderedMap"
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
'* Class:  GenericOrderedMap
'*
'* Purpose: OrderedMap offers GenericLinkedMap syntax with ordering.  Objects
'*          added or inserted in an OrderedLookupTable must have both a key and an index, and
'*          can be retrieved by either. A key cannot be null, but a value can be, if the value type is a reference type.
'*
'*
'*======================================================================================================================

'@Exposed
Option Explicit
'@Folder "<T>Enumerable"

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
Public Property Get Map() As GenericLinkedMap: Set Map = Me: End Property
Public Property Get Comparer() As IGenericComparer: Set Comparer = myclass.Comparer: End Property
Public Property Get First() As Long: First = myclass.KeyList.First: End Property
Public Property Get Last() As Long: Last = myclass.KeyList.Last: End Property

Public Property Get GetKey(ByVal Index As Long) As IGeneric
    
    If (Index >= myclass.KeyList.Elements.Count) Then _
        Call System.Throw(Me, "Public Property Get GetKey").IndexOutofRange("Index not valid")
    
    Set GetKey = myclass.KeyList.ElementAt(Index)

End Property

Public Property Get Item(ByVal Key As IGeneric) As IGeneric: Set Item = myclass.LookupTable.Item(Key): End Property
Attribute Item.VB_UserMemId = 0
Public Property Set Item(ByVal Key As IGeneric, ByVal Item As IGeneric)
    
    If (myclass.LookupTable.Add(Key, Item, Overwrite:=True)) Then _
        Call myclass.KeyList.Add(Key)
        
End Property
Public Property Get ElementAt(ByVal Index As Long): Set ElementAt = myclass.LookupTable.Item(Me.GetKey(Index)): End Property
Public Property Set ElementAt(ByVal Index As Long, ByVal Item As IGeneric): Set myclass.LookupTable.Item(Me.GetKey(Index)) = Item: End Property

Public Function Build(Optional ByVal Capacity As Long = DEFAULT_CAPACITY, Optional ByVal Comparer As IGenericComparer) As GenericOrderedMap
    
    Dim newMap As Member
    
    If Comparer Is Nothing Then _
        Set Comparer = IGenericComparer
    Set newMap.Comparer = Comparer
    
    Set newMap.KeyList = GenericOrderedList.Build(Capacity)
    Set newMap.LookupTable = GenericLinkedMap.Build(Capacity * 2, newMap.Comparer)
    
    Set Build = System.CreateInstance(New GenericOrderedMap, VarPtr(newMap), LenB(newMap))

End Function

Public Function GetEntry(ByVal Index As Long) As GenericPair

    Dim Key As IGeneric
    Set Key = Me.GetKey(Index)
    Set GetEntry = GenericPair.Build(Key, myclass.LookupTable.Item(Key))
    
End Function

Public Function GetKeys() As GenericArray: Set GetKeys = myclass.KeyList.Elements.ToArray: End Function
Public Function GetValues() As GenericArray: Set GetValues = myclass.LookupTable.GetValues: End Function

Public Sub Add(ByVal Key As IGeneric, ByVal Item As IGeneric)
    
    If (myclass.LookupTable.Add(Key, Item, Overwrite:=False)) Then _
        Call myclass.KeyList.Add(Key)
   
End Sub

Public Sub AddAll(ByVal Map As GenericLinkedMap)
    
    Dim p As GenericPair
  
    With Map.Elements.Iterator()
        Do While .HasNext(p): Call Me.Add(p.Key, p.Value): Loop
    End With
    
End Sub

Public Function ContainsKey(ByVal Key As IGeneric) As Boolean: ContainsKey = myclass.LookupTable.ContainsKey(Key): End Function
Public Function ContainsValue(ByVal Value As IGeneric) As Boolean: ContainsValue = myclass.LookupTable.ContainsValue(Value): End Function

Public Function TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: TryGetValue = myclass.LookupTable.TryGetValue(Key, Value): End Function
Public Function IndexOfKey(ByVal Key As IGeneric) As Long

    If (Key Is Nothing) Then _
        Call System.Throw(Me, "Public Function IndexOfKey").IsNothing("(Key Is Nothing)")
    
    IndexOfKey = myclass.KeyList.IndexOf(Element:=Key, Comparer:=myclass.Comparer)
    
End Function
'Use equality
Public Function IndexOfValue(ByVal Value As IGeneric) As Long
    
    Dim i As Long
    
    For i = myclass.KeyList.First To myclass.KeyList.Last
        If (myclass.LookupTable.Item(myclass.KeyList.ElementAt(i)).Equals(Value)) Then _
            IndexOfValue = i: Exit Function
    Next
   
    IndexOfValue = -1

End Function

Public Sub RemoveAt(ByVal Index As Long)

    If (myclass.LookupTable.Remove(Me.GetKey(Index))) Then _
        Call myclass.KeyList.RemoveAt(Index)
        
End Sub

Public Function Remove(ByVal Key As IGeneric) As Boolean
    
    Dim i As Long
    i = Me.IndexOfKey(Key:=Key)
    
    If (i >= 0) Then
        If (myclass.LookupTable.Remove(myclass.KeyList.ElementAt(i))) Then _
            Call myclass.KeyList.RemoveAt(i): Remove = True
        
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

Private Property Get IGenericReadOnlyList_ElementAt(ByVal Index As Long) As IGeneric
    
    If (Index >= myclass.KeyList.Elements.Count) Then _
        Call System.Throw(Me, "Private Property Get IGenericReadOnlyList_ElementAt").IndexOutofRange("(Index >= myclass.KeyList.Count)")
    
    Set IGenericReadOnlyList_ElementAt = Me.GetEntry(Index)
    
End Property

Private Function IGenericReadOnlyList_IndexOf(ByVal Element As IGeneric) As Long
  
    If Not (Element Is Nothing) Then _
        If myclass.LookupTable.ContainsKey(Element) Then _
            IGenericReadOnlyList_IndexOf = myclass.KeyList.IndexOf(Element): Exit Function
  
    IGenericReadOnlyList_IndexOf = -1
    
End Function

'IGenericMap
Private Property Get IGenericMap_Elements() As IGenericCollection
    Set IGenericMap_Elements = Me
End Property

Private Property Set IGenericMap_Item(ByVal Key As IGeneric, ByVal Item As IGeneric)

End Property

Private Property Get IGenericMap_Item(ByVal Key As IGeneric) As IGeneric

End Property

Private Property Get IGenericMap_Comparer() As IGenericComparer

End Property

Private Function IGenericMap_GetKeys() As GenericArray
    Set IGenericMap_GetKeys = Me.GetKeys
End Function

Private Function IGenericMap_GetValues() As GenericArray
    Set IGenericMap_GetValues = Me.GetValues
End Function

Private Function IGenericMap_ContainsKey(ByVal Key As IGeneric) As Boolean
    
End Function

Private Function IGenericMap_ContainsValue(ByVal Value As IGeneric) As Boolean

End Function

Private Function IGenericMap_Add(ByVal Key As IGeneric, ByVal Value As IGeneric, Optional ByVal Overwrite As Boolean) As Boolean

End Function

Private Sub IGenericMap_AddAll(ByVal Map As IGenericMap, Optional ByVal Overwrite As Boolean)

End Sub

Private Function IGenericMap_Remove(ByVal Key As IGeneric) As Boolean

End Function

Private Function IGenericMap_RemoveAll(ByVal Collection As IGenericCollection) As Long

    

End Function

Private Function IGenericMap_TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean

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
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.ContainsKey(Item): End Function
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = myclass.LookupTable.Elements.Iterator: End Function
Private Function IGenericCollection_Copy() As IGenericCollection
    
    Dim c As Member
    
    Set c.Comparer = myclass.Comparer
    Set c.LookupTable = myclass.LookupTable.Elements.Copy
    Set c.KeyList = myclass.KeyList.Elements.Copy
    
    Set IGenericCollection_Copy = System.CreateInstance(New GenericOrderedMap, VarPtr(c), LenB(c))
    
End Function

