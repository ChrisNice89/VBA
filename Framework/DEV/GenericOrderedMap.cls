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
    index As Long
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

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get BaseIndex() As Long: BaseIndex = myclass.Internal.BaseIndex: End Property
Public Property Get Entry(ByVal index As Long) As GenericPair: Set Entry = myclass.Internal(index): End Property
Public Property Get Count() As Long: Count = myclass.Internal.Count: End Property
Public Property Get GetKey(ByVal index As Long) As IGeneric
    
    If (index > myclass.Internal.Count) Then _
        Call Skynet.Throw(Me, "Public Property Get GetKey").IndexOutofRange("Index not valid")
    
    Set GetKey = Me.Entry(index).Key

End Property

Public Property Get item(ByVal Key As IGeneric) As IGeneric: Set item = myclass.HashTable(Key): End Property
Attribute item.VB_UserMemId = 0
Public Property Set item(ByVal Key As IGeneric, ByVal item As IGeneric)
    With myclass
        If (.HashTable.ContainsKey(Key)) Then _
            Set .Internal(IndexOfKey(Key)) = GenericPair(Key, item)

        Set .HashTable(Key) = item
    End With
End Property
Public Property Get ElementAt(ByVal index As Long): Set ElementAt = Me.Entry(index).Value: End Property
Public Property Set ElementAt(ByVal index As Long, ByVal item As IGeneric)
    
    Dim Key As IGeneric
    
    If (index > myclass.Internal.Count) Then _
        Call Skynet.Throw(Me, "Public Property Set ElementAt").IndexOutofRange("Index not valid")
    
    Set Key = Me.Entry(index).Key
    Set myclass.HashTable(Key) = item
    Set myclass.Internal(index) = GenericPair(Key, item)
    
End Property

Public Function Build(Optional ByVal Capacity As Long = DEFAULT_CAPACITY) As GenericOrderedMap
    
    If (Me Is GenericOrderedMap) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption
    
    With myclass
        Set .Internal = GenericList.Build(Capacity)
        Set .HashTable = GenericMap.Build(Capacity * 2) 'GenericTree.Build
    End With

    Set Build = Skynet.CreateInstance(New GenericOrderedMap, Virtual.Pointer, Virtual.Size)

End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
End Sub

Public Function GetKeys() As GenericArray

    ReDim Keys(myclass.Internal.BaseIndex To myclass.Internal.Count) As IGeneric

    Dim i As Long
    
    With Me
        For i = LBound(Keys) To UBound(Keys)
            Set Keys(i) = .Entry(i).Key
        Next
    End With
    
    Set GetKeys = GenericArray.BuildFrom(Keys)
    
End Function

Public Function GetValues() As GenericArray

    ReDim Values(myclass.Internal.BaseIndex To myclass.Internal.Count) As IGeneric

    Dim i As Long
    
    With Me
        For i = LBound(Values) To UBound(Values)
            Set Values(i) = Me.Entry(i).Value
        Next
    End With
    
    Set GetValues = GenericArray.BuildFrom(Values)

End Function

Public Sub Add(ByVal Key As IGeneric, ByVal item As IGeneric)
    With myclass
        Call .HashTable.Add(Key, item)
        Call .Internal.Add(GenericPair(Key, item))
    End With
End Sub

Public Sub AddAll(ByVal Map As IGenericDictionary)
    
    Dim P As GenericPair
  
    With Map.IteratorOf(PairData) ' Copy
        Do While .HasNext(P): Call Me.Add(P.Key, P.Value): Loop
    End With
    
End Sub

Public Sub AddPair(ByVal Pair As GenericPair)
    
    With myclass
        Call .HashTable.Add(Pair.Key, Pair.Value)
        Call .Internal.Add(Skynet.Clone(Pair))
    End With
    
End Sub
Public Function ContainsKey(ByVal Key As IGeneric) As Boolean: ContainsKey = myclass.HashTable.ContainsKey(Key): End Function
Public Function TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: TryGetValue = myclass.HashTable.TryGetValue(Key, Value): End Function
Public Function IndexOfKey(ByVal Key As IGeneric) As Long
    
    Dim i As Long
    With myclass
        For i = .Internal.BaseIndex To .Internal.Count
            If (Key.Equals(Me.Entry(i).Key)) Then _
                IndexOfKey = i: Exit Function
        Next
    End With
    IndexOfKey = -1

End Function

Public Function IndexOfValue(ByVal Value As IGeneric) As Long
    
    Dim i As Long
    With myclass
        For i = .Internal.BaseIndex To .Internal.Count
            If (Value.Equals(Me.Entry(i).Value)) Then _
                IndexOfValue = i: Exit Function
        Next
    End With
    
    IndexOfValue = -1

End Function

Public Sub Insert(ByVal index As Long, ByVal Key As IGeneric, ByVal Value As IGeneric)
    
    If (index > myclass.Internal.Count) Or (index < 0) Then _
        Call Skynet.Throw(Me, "Public Sub Insert").IndexOutofRange("(Index > myClass.Internal.Count) Or (Index < 0)")
        
    With myclass
        Call .HashTable.Add(Key, Value)
        Call .Internal.Insert(index, GenericPair(Key, Value))
    End With
    
End Sub

Public Sub Clear()
    Call myclass.Internal.Clear
    Call myclass.HashTable.Clear
End Sub

Public Sub RemoveAt(ByVal index As Long)
    
    If (index < 0) Then _
        Exit Sub
        
    With myclass
        Call .HashTable.Remove(Me.Entry(index).Key)
        Call .Internal.RemoveAt(index)
    End With

End Sub

Public Sub Remove(ByVal Key As IGeneric)

    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Sub Remove").KeyIsNothing
        
    Call Me.RemoveAt(Me.IndexOfKey(Key))
End Sub

Private Function CopyOf(ByVal t As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long, ByVal Caller As String) As GenericArray
    
    If (GArray Is Nothing) Then _
        Call Skynet.Throw(Me, Caller).NullPointer("GArray Is Nothing")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call Skynet.Throw(Me, Caller).IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < myclass.Internal.Count) Then _
        Call Skynet.Throw(Me, Caller).InvalidInput("(GArray.Length - ArrayIndex + GArray.LowerBound < myClass.Internal.Count)")
    
    Dim i As Long
    
    Select Case True
        Case t = KeyData
            For i = 1 To myclass.Internal.Count
                Set GArray(ArrayIndex) = Me.Entry(i).Key: Call InterlockedIncrement(ArrayIndex)
            Next
        Case t = ValueData
            For i = 1 To myclass.Internal.Count
                Set GArray(ArrayIndex) = Me.Entry(i).Value: Call InterlockedIncrement(ArrayIndex)
            Next
        Case Else
            Call myclass.Internal.CopyTo(index:=GenericArray.LowerBound, GArray:=GArray, ArrayIndex:=ArrayIndex, Elements:=myclass.Internal.Count)
    End Select
    
    Set CopyOf = GArray
    
End Function

Public Function IteratorOf(ByVal t As DictionaryData, Optional ByVal index As Long, Optional ByVal Count As Long) As IGenericIterator
    
    Dim c As Member
    
    If index = 0 Then _
        index = myclass.Internal.BaseIndex
        
    If Count = 0 Then _
        Count = myclass.Internal.Count
    
    If (index < myclass.Internal.BaseIndex) Or (index > myclass.Internal.Count + myclass.Internal.BaseIndex) Then _
        Call Skynet.Throw(Me, "Private Sub LoadIterator").IndexOutofRange("(Index < myClass.Internal.BaseIndex) Or (Index > myClass.Internal.Count + myClass.Internal.BaseIndex)")

    If (Count < 0) Or (Count > myclass.Internal.Count - index + myclass.Internal.BaseIndex) Then _
        Call Skynet.Throw(Me, "Public Function Iterator").ArgOutOfRange("(Count < 0) Or (Count > myClass.Internal.Count - Index + myClass.Internal.BaseIndex)")
    
    If t = PairData Then
        Set IteratorOf = myclass.Internal.Iterator(index:=index, Count:=Count)
    Else
        Call LoadIterator(c.Iterator, t, myclass.Internal, index, Count)
        Set IteratorOf = Skynet.CreateInstance(New GenericOrderedMap, VarPtr(c), LenB(c))
    End If
    
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
        .index = .StartIndex - 1
        .EndIndex = .index + .Count
        .HasLoaded = True
    End With
    
End Sub

'IGeneric
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GenericOrderedMap: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = Virtual.Pointer: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = Virtual.Size: End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericOrderedMap):  End Property

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member
    Set c.HashTable = Skynet.Clone(myclass.HashTable)
    Set c.Internal = Skynet.Clone(myclass.Internal)
    
    Set IGeneric_Clone = Skynet.CreateInstance(New GenericOrderedMap, VarPtr(c), Virtual.Size)
   
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
Private Property Get IGenericReadOnlyList_BaseIndex() As Long: IGenericReadOnlyList_BaseIndex = myclass.Internal.BaseIndex: End Property
Private Property Get IGenericReadOnlyList_Count() As Long: IGenericReadOnlyList_Count = myclass.Internal.Count: End Property
Private Property Get IGenericReadOnlyList_ElementAt(ByVal index As Long) As IGeneric
    
    Dim P As GenericPair
    
    If (index > myclass.Internal.Count) Then _
        Call Skynet.Throw(Me, "Private Property Get IGenericReadOnlyList_ElementAt").IndexOutofRange("Index not valid")
    
    Set P = myclass.Internal(index)
    
    Set IGenericReadOnlyList_ElementAt = P.Value

End Property

Private Sub IGenericReadOnlyList_Clear(): Call Me.Clear: End Sub
Private Sub IGenericReadOnlyList_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(t:=ValueData, GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="Private Sub IGenericReadOnlyList_CopyTo"): End Sub
Private Function IGenericReadOnlyList_Iterator() As IGenericIterator: Set IGenericReadOnlyList_Iterator = Me.IteratorOf(t:=ValueData): End Function
Private Function IGenericReadOnlyList_IndexOf(ByVal item As IGeneric) As Long: IGenericReadOnlyList_IndexOf = Me.IndexOfValue(item): End Function

'IGenericDictionary
Private Property Set IGenericDictionary_Item(ByVal Key As IGeneric, ByVal item As IGeneric): Set Me.item(Key) = item: End Property
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
Private Function IGenericCollection_Contains(ByVal item As IGeneric) As Boolean: IGenericCollection_Contains = myclass.HashTable.ContainsKey(item): End Function
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.IteratorOf(t:=PairData): End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.Current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    Dim P As GenericPair
    
    With myclass.Iterator
        If .index = .EndIndex Then
            If .HasLoaded = False Then _
                Call Skynet.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator
            
            Set Out = Nothing
            Set .Current = Nothing
            Exit Function
        End If
        
        Set P = .Entries(InterlockedIncrement(.index))
        
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
            Call Skynet.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
                
        Call LoadIterator(myclass.Iterator, .Type, .Entries, .StartIndex, .Count)
    End With
End Sub
