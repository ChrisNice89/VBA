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
'* Purpose: OrderedMap offers GenericMap syntax with ordering.  Objects
'*          added or inserted in an OrderedHashTable must have both a key and an index, and
'*          can be retrieved by either. A key cannot be null, but a value can be, if the value type is a reference type.
'*
'*
'*======================================================================================================================

'@Exposed
Option Explicit
'@Folder "<T>Enumerable"

Implements IGeneric
Implements IGenericReadOnlyList
Implements IGenericDictionary
Implements IGenericCollection
Implements IGenericIterator

Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long

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
Private myclass As Member

Public Property Get First() As Long: First = myclass.Internal.First: End Property
Public Property Get Last() As Long: Last = myclass.Internal.Last: End Property
Public Property Get Entry(ByVal Index As Long) As GenericPair: Set Entry = myclass.Internal(Index): End Property
Public Property Get Count() As Long
    If Not myclass.Internal Is Nothing Then _
        Count = myclass.Internal.Count
End Property
Public Property Get GetKey(ByVal Index As Long) As IGeneric
    
    If (Index > myclass.Internal.Count) Then _
        Call System.Throw(Me, "Public Property Get GetKey").IndexOutofRange("Index not valid")
    
    Set GetKey = Me.Entry(Index).Key

End Property

Public Property Get Item(ByVal Key As IGeneric) As IGeneric: Set Item = myclass.HashTable(Key): End Property
Attribute Item.VB_UserMemId = 0
Public Property Set Item(ByVal Key As IGeneric, ByVal Item As IGeneric)
    With myclass
        If (.HashTable.ContainsKey(Key) = False) Then _
                Call Me.Add(Key, Item): Exit Property
                
        Set .HashTable(Key) = Item
        Set .Internal(IndexOfKey(Key)) = GenericPair(Key, Item)
        
    End With
End Property
Public Property Get ElementAt(ByVal Index As Long): Set ElementAt = Me.Entry(Index).Value: End Property
Public Property Set ElementAt(ByVal Index As Long, ByVal Item As IGeneric)
    
    Dim Key As IGeneric
    
    If (Index > myclass.Internal.Count) Then _
        Call System.Throw(Me, "Public Property Set ElementAt").IndexOutofRange("Index not valid")
    
    Set Key = Me.Entry(Index).Key
    Set myclass.HashTable(Key) = Item
    Set myclass.Internal(Index) = GenericPair(Key, Item)
    
End Property

Public Function Build(Optional ByVal Capacity As Long = DEFAULT_CAPACITY) As GenericOrderedMap
    
    Dim newMap As Member
    
    Set newMap.Internal = GenericList.Build(Capacity)
    Set newMap.HashTable = GenericLinkedMap.Build(Capacity * 2)
   
    Set Build = System.CreateInstance(New GenericOrderedMap, VarPtr(newMap), LenB(newMap))

End Function

Public Function GetKeys() As GenericArray

    Dim Keys As GenericArray
    Dim i As Long, Index As Long
    
    Set Keys = GenericArray.Build(Me.Count)
    
    With Keys
        For i = myclass.Internal.First To myclass.Internal.Last
            Set .ElementAt(i) = Me.Entry(i).Key
        Next
    End With
    
    Set GetKeys = Keys
    
End Function

Public Function GetValues() As GenericArray

    Dim Values As GenericArray
    Dim i As Long, Index As Long
    
    Set Values = GenericArray.Build(Me.Count)
    
    With Values
        For i = myclass.Internal.First To myclass.Internal.Last
            Set .ElementAt(i) = Me.Entry(i).Value
        Next
    End With
    
    Set GetValues = Values

End Function

Public Sub Add(ByVal Key As IGeneric, ByVal Item As IGeneric)
    
    Call myclass.HashTable.Add(Key, Item)
    Call myclass.Internal.Add(GenericPair(Key, Item))
   
End Sub

Public Sub AddAll(ByVal Map As IGenericDictionary)
    
    Dim P As GenericPair
  
    With Map.IteratorOf(PairData)
        Do While .HasNext(P): Call Me.AddPair(P): Loop
    End With
    
End Sub

Public Sub AddPair(ByVal pair As GenericPair)
    
    Call myclass.HashTable.Add(pair.Key, pair.Value)
    Call myclass.Internal.Add(pair)
   
End Sub
Public Function ContainsKey(ByVal Key As IGeneric) As Boolean: ContainsKey = myclass.HashTable.ContainsKey(Key): End Function
Public Function TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: TryGetValue = myclass.HashTable.TryGetValue(Key, Value): End Function
Public Function IndexOfKey(ByVal Key As IGeneric) As Long
    
    Dim i As Long
    With myclass
        For i = .Internal.First To .Internal.Count
            If (Key.Equals(Me.Entry(i).Key)) Then _
                IndexOfKey = i: Exit Function
        Next
    End With
    IndexOfKey = -1

End Function

Public Function IndexOfValue(ByVal Value As IGeneric) As Long
    
    Dim i As Long
    With myclass
        For i = .Internal.First To .Internal.Count
            If (Value.Equals(Me.Entry(i).Value)) Then _
                IndexOfValue = i: Exit Function
        Next
    End With
    
    IndexOfValue = -1

End Function

Public Sub Insert(ByVal Index As Long, ByVal Key As IGeneric, ByVal Value As IGeneric)
    
    If (Index > myclass.Internal.Count) Or (Index < 0) Then _
        Call System.Throw(Me, "Public Sub Insert").IndexOutofRange("(Index > myClass.Internal.Count) Or (Index < 0)")
        
    With myclass
        Call .HashTable.Add(Key, Value)
        Call .Internal.Insert(Index, GenericPair(Key, Value))
    End With
    
End Sub

Public Sub Clear()
    
    With myclass
        Call .Internal.Clear
        Call .HashTable.Clear
    End With
    
End Sub

Public Sub RemoveAt(ByVal Index As Long)
    
    If (Index < 0) Then _
        Exit Sub
        
    With myclass
        Call .HashTable.Remove(Me.Entry(Index).Key)
        Call .Internal.RemoveAt(Index)
    End With

End Sub

Public Sub Remove(ByVal Key As IGeneric)

    If (Key Is Nothing) Then _
        Call System.Throw(Me, "Public Sub Remove").IsNothing("(Key Is Nothing)")
        
    Call Me.RemoveAt(Me.IndexOfKey(Key))
    
End Sub

Private Function CopyOf(ByVal t As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long, ByVal Caller As String) As GenericArray
    
    If (GArray Is Nothing) Then _
        Call System.Throw(Me, Caller).IsNothing("(GArray Is Nothing)")
    
    If (ArrayIndex < GArray.LowerBound) Then _
        Call System.Throw(Me, Caller).IndexOutofRange("(ArrayIndex < GArray.LowerBound)")

    If (GArray.Length - ArrayIndex < myclass.Internal.Count) Then _
        Call System.Throw(Me, Caller).InvalidInput("(GArray.Length - ArrayIndex < myclass.Internal.Count)")
    
    Dim i As Long
    
    Select Case True
        Case t = KeyData
            ArrayIndex = ArrayIndex - 1
            For i = 1 To myclass.Internal.Count
                Set GArray(InterlockedIncrement(ArrayIndex)) = Me.Entry(i).Key
            Next
        Case t = ValueData
            ArrayIndex = ArrayIndex - 1
            For i = 1 To myclass.Internal.Count
                Set GArray(InterlockedIncrement(ArrayIndex)) = Me.Entry(i).Value
            Next
        Case Else
            Call myclass.Internal.CopyTo(Index:=GenericArray.LowerBound, GArray:=GArray, ArrayIndex:=ArrayIndex, Count:=myclass.Internal.Count)
    End Select
    
    Set CopyOf = GArray
    
End Function

Public Function IteratorOf(ByVal t As DictionaryData, Optional ByVal Index As Long, Optional ByVal Count As Long) As IGenericIterator
    
    Dim c As Member
    
    If (Index < 0) Then _
        Call System.Throw(Me, "Private Sub LoadIterator").IndexOutofRange("(Index < 0)")
    
    If Count = 0 Then
        Count = myclass.Internal.Count
    ElseIf (Count < 0) Then
        Call System.Throw(Me, "Public Function Iterator").ArgOutOfRange("(Count < 0)")
    End If
    
    If (Count > myclass.Internal.Count - Index) Then _
        Call System.Throw(Me, "Public Function Iterator").ArgOutOfRange("(Count > myclass.Internal.Count - Index)")

     Call LoadIterator(c.Iterator, t, myclass.Internal, Index, Count)
     Set IteratorOf = System.CreateInstance(New GenericOrderedMap, VarPtr(c), LenB(c))
    
End Function

Private Sub LoadIterator( _
                            ByRef Iterator As Iterator, _
                            ByVal t As DictionaryData, _
                            ByVal Entries As GenericList, _
                            ByVal StartIndex As Long, _
                            ByVal Count As Long)
 
    With Iterator
        .Type = t
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
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericOrderedMap: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GenericOrderedMap: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericOrderedMap):  End Property

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member
    Set c.HashTable = System.Clone(myclass.HashTable)
    Set c.Internal = System.Clone(myclass.Internal)
    
    Set IGeneric_Clone = System.CreateInstance(New GenericOrderedMap, VarPtr(c), LenB(c))
   
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
                            "Elements: " & Me.Count
    End With
End Function

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_First() As Long: IGenericReadOnlyList_First = myclass.Internal.First: End Property
Private Property Get IGenericReadOnlyList_Last() As Long: IGenericReadOnlyList_Last = myclass.Internal.Last: End Property
Private Property Get IGenericReadOnlyList_Count() As Long: IGenericReadOnlyList_Count = Me.Count: End Property
Private Property Get IGenericReadOnlyList_ElementAt(ByVal Index As Long) As IGeneric
    
    Dim P As GenericPair
    
    If (Index >= myclass.Internal.Count) Then _
        Call System.Throw(Me, "Private Property Get IGenericReadOnlyList_ElementAt").IndexOutofRange("(Index >= myclass.Size)")
    
    Set P = myclass.Internal(Index)
    
    Set IGenericReadOnlyList_ElementAt = P.Value

End Property

Private Sub IGenericReadOnlyList_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(t:=ValueData, GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="Private Sub IGenericReadOnlyList_CopyTo"): End Sub
Private Function IGenericReadOnlyList_Iterator() As IGenericIterator: Set IGenericReadOnlyList_Iterator = Me.IteratorOf(t:=ValueData): End Function
Private Function IGenericReadOnlyList_IndexOf(ByVal Item As IGeneric) As Long: IGenericReadOnlyList_IndexOf = Me.IndexOfValue(Item): End Function

'IGenericDictionary
Private Property Set IGenericDictionary_Item(ByVal Key As IGeneric, ByVal Item As IGeneric): Set Me.Item(Key) = Item: End Property
Private Property Get IGenericDictionary_Item(ByVal Key As IGeneric) As IGeneric: Set IGenericDictionary_Item = myclass.HashTable(Key): End Property
Private Property Get IGenericDictionary_Count() As Long: IGenericDictionary_Count = myclass.HashTable.Count: End Property
Private Function IGenericDictionary_TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean:  IGenericDictionary_TryGetValue = myclass.HashTable.TryGetValue(Key, Value): End Function
Private Function IGenericDictionary_ContainsKey(ByVal Key As IGeneric) As Boolean: IGenericDictionary_ContainsKey = myclass.HashTable.ContainsKey(Key): End Function
Private Function IGenericDictionary_ContainsValue(ByVal Value As IGeneric) As Boolean: IGenericDictionary_ContainsValue = Me.IndexOfValue(Value) > -1: End Function
Private Sub IGenericDictionary_Add(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value): End Sub
Private Sub IGenericDictionary_Remove(ByVal Key As IGeneric): Call Me.Remove(Key): End Sub
Private Sub IGenericDictionary_CopyOf(ByVal t As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(t:=t, GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="Private Sub IGenericDictionary_CopyOf"): End Sub
Private Sub IGenericDictionary_Clear(): Call Me.Clear: End Sub
Private Function IGenericDictionary_IteratorOf(ByVal t As DictionaryData) As IGenericIterator: Set IGenericDictionary_IteratorOf = Me.IteratorOf(t:=t): End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = False: End Property
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(t:=PairData, GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="Private Sub IGenericCollection_CopyTo"): End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = myclass.Internal.ToArray: End Function
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = myclass.HashTable.ContainsKey(Item): End Function
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.IteratorOf(t:=PairData): End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.Current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    Dim P As GenericPair
    
    With myclass.Iterator
       
        If .Index = .EndIndex Then
            If .HasLoaded = False Then _
                Call System.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator
            
            Set .Current = Nothing
            IGenericIterator_HasNext = False
            Exit Function
        End If
            
        Set P = .Entries(InterlockedIncrement(.Index))
        
        If .Type = KeyData Then
            Set .Current = P.Key
        
        ElseIf .Type = ValueData Then
            Set .Current = P.Value
        
        Else
            Set .Current = P
        End If
        
        Set Out = .Current
        IGenericIterator_HasNext = True
        
    End With

End Function

Private Sub IGenericIterator_Reset()
    With myclass.Iterator
        If .HasLoaded = False Then _
            Call System.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
                
        Call LoadIterator(myclass.Iterator, .Type, .Entries, .StartIndex, .Count)
    End With
End Sub
