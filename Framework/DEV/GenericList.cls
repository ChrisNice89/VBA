VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericList"
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
'* Class:  GenericList
'*
'* Purpose: GeneriList is a resizable array
'*          i.e. GenericList grows dynamically as the elements are added to it.
'*          If the Size of the current elements (including the new element to be added to the GenericList)
'*          is greater than the maximum Size of the array then increase the Size of array.
'*          But the Size of the array can not be increased dynamically.
'*          So, what happens internally is, a new Array is created and
'*          the old array is copied into the new array.
'*
'*
'*======================================================================================================================

'@Folder("<T>Enumerable")

Option Explicit
'@PredeclaredId

Implements IPrototype
Implements IGeneric
Implements IGenericReadOnlyList
Implements IGenericCollection
Implements IGenericIterator

Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long

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
    Size As Long
    IsReadOnly As Boolean
    SortOrder As SortOrder
    Iterator As Iterator
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Count() As Long: Count = myClass.Size: End Property
Public Property Get BaseIndex() As Long: BaseIndex = myClass.Items.LowerBound: End Property
Public Property Get SortOrder() As SortOrder: SortOrder = myClass.SortOrder: End Property
Public Property Let SortOrder(ByVal Order As SortOrder):    myClass.SortOrder = Order: End Property
Public Property Get Last() As IGeneric: Set Last = myClass.Items(myClass.Size): End Property
Public Property Set ElementAt(ByVal Index As Long, ByVal Item As IGeneric)

    If (Index > myClass.Size) Then _
        Call Skynet.Throw(Me, "Public Property Get ElementAt").IndexOutofRange("Index not valid")
    
    Set myClass.Items(Index) = Item

End Property

Public Property Get ElementAt(ByVal Index As Long) As IGeneric
Attribute ElementAt.VB_UserMemId = 0

    If (Index > myClass.Size) Then _
        Call Skynet.Throw(Me, "Public Property Get ElementAt").IndexOutofRange("Index not valid")
    
    Set ElementAt = myClass.Items(Index)

End Property

Public Property Get Capacity() As Long: Capacity = myClass.Items.Length: End Property
Public Property Let Capacity(ByVal Value As Long)
    
    If (Value < myClass.Size) Then _
        Call Skynet.Throw(Me, "Public Property Let Capacity").ArgOutOfRange("Value < myClass.Size")
    
    If (myClass.Items Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Property Let Capacity").NullPointer("myClass.Items Is Nothing")
    
    If Not (Value = myClass.Items.Length) Then
        If (Value > 0) Then
            Dim newItems As GenericArray
            Set newItems = GenericArray.Build(Value)
            If (myClass.Size > 0) Then _
                Call Skynet.CopyData(dst:=newItems.AdressOf(Element:=1), src:=myClass.Items.AdressOf(Element:=1), Bytes:=myClass.Items.TotalSize, ClearSource:=True)

            Set myClass.Items = newItems
        Else
            Call myClass.Items.Clear
        End If
    End If
End Property

Private Sub EnsureCapacity(ByVal Min As Long)
    
    Dim newCapacity As Long
    
    With myClass.Items
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

Public Function Build( _
                        Optional ByVal Capacity As Long = DEFAULT_CAPACITY, _
                        Optional ByVal collection As IGenericCollection) As GenericList
    
    If (Me Is GenericList) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption
    
    If (Capacity < 0) Then _
        Call Skynet.Throw(Me, "Build").ArgOutOfRange("Capacity < 0")
    
    With myClass
        .SortOrder = Skynet.SortOrder
        If (collection Is Nothing) Then
            Set .Items = GenericArray.Build(Capacity)
        Else
            Set .Items = GenericArray.Build(collection.Count)
            Call collection.CopyTo(GArray:=.Items, ArrayIndex:=.Items.LowerBound)
            .Size = collection.Count
        End If
    End With
    
    Set Build = Skynet.CreateInstance(New GenericList, Virtual.Pointer, Virtual.Size)
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Virtual.Size = LenB(myClass)
End Sub

Private Sub Class_Terminate()
    Call Skynet.Dispose(Me)
End Sub

Public Function AsReadOnly() As IGenericReadOnlyList
    Set AsReadOnly = Me 'GenericReadOnlyList.Build(Internal:=myClass.Items, Size:=myClass.Size)
End Function

Public Function Iterator(Optional ByVal Index As Long, Optional ByVal Count As Long) As IGenericIterator
    
    Dim c As Member
    
    If Index = 0 Then _
        Index = myClass.Items.LowerBound
        
    If Count = 0 Then _
        Count = myClass.Size
    
    If (Index < myClass.Items.LowerBound) Or (Index > myClass.Size + myClass.Items.LowerBound) Then _
        Call Skynet.Throw(Me, "Private Sub LoadIterator").IndexOutofRange("(Index < myClass.Items.LowerBound) Or (Index > myClass.Size + myClass.Items.LowerBound)")

    If (Count < 0) Or (Count > myClass.Size - Index + myClass.Items.LowerBound) Then _
        Call Skynet.Throw(Me, "Public Function Iterator").ArgOutOfRange("(Count < 0) Or (Count > myClass.Size - Index + myClass.Items.LowerBound)")
        
    'Call LoadIterator(C.Iterator, myClass.Items, Index, Count)
    'Set Iterator = Skynet.CreateInstance(New GenericList, VarPtr(C), LenB(C))
    Set Iterator = myClass.Items.Iterator(StartIndex:=Index, Count:=Count)

End Function

Public Function Add(ByVal Item As IGeneric) As Long

    If (myClass.Size = Me.Capacity) Then _
        Call EnsureCapacity(myClass.Size + 1)
    
    Set myClass.Items(InterlockedIncrement(myClass.Size)) = Item
    Add = myClass.Size

End Function

Public Sub AddAll(ByVal collection As IGenericIterator)
    Call Me.InsertAll(Index:=myClass.Size + myClass.Items.LowerBound, Enumerable:=collection)
End Sub

Public Sub Insert(ByVal Index As Long, ByVal Item As IGeneric)
    
    If (Index > myClass.Size + myClass.Items.LowerBound) Then _
        Call Skynet.Throw(Me, "Public Sub Insert").IndexOutofRange("Index > myClass.Size")
    
    If (myClass.Size = myClass.Items.Length) Then _
        Call EnsureCapacity(myClass.Size + 1)
        
    With myClass.Items
        If (Index < myClass.Size + .LowerBound) Then _
            Call Skynet.CopyData(dst:=ByVal .AdressOf(Element:=Index + 1), src:=ByVal .AdressOf(Element:=Index), Bytes:=.SizeOf(myClass.Size - Index + .LowerBound)): _
            Call Skynet.ClearData(dst:=ByVal .AdressOf(Element:=Index), Bytes:=.SizeOf(1))

         Set .ElementAt(Index) = Item: Call InterlockedIncrement(myClass.Size)
    End With

End Sub

Public Sub InsertAll(ByVal Index As Long, ByVal Enumerable As IGenericIterator)
    
    Dim c As IGenericCollection
    Dim n As Long
    Dim Item As IGeneric
    
    If (Enumerable Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Sub InsertAll").NullPointer("Collection Is Nothing")
    
    If (Index > myClass.Size + myClass.Items.LowerBound) Then _
        Call Skynet.Throw(Me, "InsertAll").IndexOutofRange("Index > myClass.Size")
    
    If TypeOf Enumerable Is IGenericCollection Then _
        Set c = Enumerable: n = c.Count
    
    If (n > 0) Then
        Call EnsureCapacity(myClass.Size + c.Count)
    
        With myClass.Items
            If (Index < myClass.Size + .LowerBound) Then _
                Call Skynet.CopyData(dst:=ByVal .AdressOf(Element:=Index + c.Count), src:=ByVal .AdressOf(Element:=Index), Bytes:=.SizeOf(myClass.Size - Index + .LowerBound)): _
                Call Skynet.ClearData(dst:=ByVal .AdressOf(Element:=Index), Bytes:=.SizeOf(c.Count))

        End With
      
        Call c.CopyTo(GArray:=myClass.Items, ArrayIndex:=Index)
        myClass.Size = myClass.Size + c.Count
    Else
        With Enumerable
            Index = Index - 1
            Do While .HasNext(Item): Call Me.Insert(InterlockedIncrement(Index), Item): Loop
            Call .Reset
        End With
    End If
    
End Sub

Public Function Remove(ByVal Item As IGeneric) As Boolean
    Dim i As Long
    i = Me.IndexOf(Item:=Item)
    
    If (i >= 0) Then _
        Call Me.RemoveAt(i): Remove = True: Exit Function
    
End Function

Public Sub RemoveAt(ByVal Index As Long)

    If (Index > myClass.Size) Then _
        Call Skynet.Throw(Me, "Public Sub RemoveAt").IndexOutofRange("Index > myClass.Size")
    
    InterlockedDecrement myClass.Size
    
    With myClass.Items
        Set .ElementAt(Index) = Nothing
        
        If (Index < myClass.Size + .LowerBound) Then _
            Call Skynet.CopyData(dst:=ByVal .AdressOf(Element:=Index), src:=ByVal .AdressOf(Element:=Index + 1), Bytes:=.SizeOf(myClass.Size - Index + .LowerBound)): _
            Call Skynet.ClearData(dst:=ByVal .AdressOf(Element:=myClass.Size), Bytes:=.SizeOf(1))
            
    End With
    
End Sub

Public Sub RemoveRange(ByVal Index As Long, ByVal Count As Long)

    If (Index > myClass.Size) Then _
        Call Skynet.Throw(Me, "Public Sub RemoveAt").ArgOutOfRange("Index > myClass.Size")
  
    
End Sub

Public Function GetRange(ByVal Index As Long, ByVal Count As Long) As GenericList
    
    Dim c As Member
    Dim LB As Long
    
    LB = GenericArray.LowerBound
    
    If (Index < LB) Then _
        Call Skynet.Throw(Me, "GetRange").IndexOutofRange("Index < 0")
    
    If (Count < LB) Then _
        Call Skynet.Throw(Me, "GetRange").ArgOutOfRange("Index < 0")
    
    If (myClass.Size - Index + LB < Count) Then _
        Call Skynet.Throw(Me, "GetRange").InvalidInput("myClass.Size - Index < Count")

    Set c.Items = GenericArray.Build(Count)
    Call Me.CopyTo(Index:=Index, GArray:=c.Items, ArrayIndex:=LB, Elements:=Count)
    c.Size = Count

    Set GetRange = Skynet.CreateInstance(New GenericList, VarPtr(c), Virtual.Size)
   
End Function

Public Sub CopyTo( _
                    ByVal Index As Long, _
                    ByVal GArray As GenericArray, _
                    ByVal ArrayIndex As Long, _
                    Optional ByVal Elements As Long = 0)

    If (Elements = 0) Then _
        Elements = myClass.Size
        
    If (GArray Is Nothing) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericCollection_CopyTo").NullPointer("GArray Is Nothing")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericCollection_CopyTo").IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < Elements) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericCollection_CopyTo").InvalidInput("(GArray.Length - Index < Elements)")
    
    'Call myClass.Items.CopyTo(GArray:=GArray, ArrayIndex:=ArrayIndex, SourceIndex:=Index, Elements:=Elements)
    Call GenericArray.Copy(SrcArray:=myClass.Items, SrcIndex:=Index, DstArray:=GArray, DstIndex:=ArrayIndex, Elements:=Elements)

End Sub

'ToArray returns a new Object array containing the contents of the List.
'This requires copying the List, which is an O(n) operation.
Public Function ToArray() As GenericArray

    Dim ga As GenericArray
    Set ga = GenericArray.Build(myClass.Size)

    Call Me.CopyTo(Index:=myClass.Items.LowerBound, GArray:=ga, ArrayIndex:=ga.LowerBound)
    Set ToArray = ga

End Function

Public Sub Sort(ByVal Order As SortOrder)
    Call myClass.Items.Sort(Order, myClass.Items.LowerBound, myClass.Size)
End Sub

Public Sub Clear():
    Call myClass.Items.Clear
    myClass.Size = 0
End Sub
Public Function Contains(Item As IGeneric) As Boolean: Contains = (Me.IndexOf(Item) > myClass.Items.LowerBound): End Function

Public Function BinarySearch( _
                                ByVal Item As IGeneric, _
                                Optional ByVal Index As Long, _
                                Optional ByVal Count As Long) As Long
    
    If (Index = 0) Then
        Index = myClass.Items.LowerBound
    Else
        If (Index < 0) Then _
            Call Skynet.Throw(Me, "Public Function BinarySearch").ArgOutOfRange("Index < 0")
    End If
    
    If (Count = 0) Then
        Count = Me.Count
    Else
        If (Count < 0) Then _
            Call Skynet.Throw(Me, "Public Function BinarySearch").ArgOutOfRange("Count < 0")
    End If
    
    If (myClass.Size - Index + myClass.Items.LowerBound < Count) Then _
        Call Skynet.Throw(Me, "Public Function BinarySearch").InvalidInput("myClass.Size - Index < Count")
    
    BinarySearch = myClass.Items.BinarySearch(Value:=Item, Index:=Index, Length:=Count, Order:=myClass.SortOrder)
    
End Function

Public Function IndexOf( _
                            ByVal Item As IGeneric, _
                            Optional ByVal Index As Long, _
                            Optional ByVal Count As Long) As Long
    
    If (Index = 0) Then
        Index = myClass.Items.LowerBound
    Else
        If (Index > myClass.Size) Or (Index < myClass.Items.LowerBound) Then _
            Call Skynet.Throw(Me, "Public Function IndexOf").IndexOutofRange("(Index > myClass.Size) Or (Index < myClass.Items.LowerBound)")
    End If
    
    If (Count = 0) Then
        Count = myClass.Size - Index + myClass.Items.LowerBound
    Else
        If (Count < 0) Or (Index > myClass.Size - Count + myClass.Items.LowerBound) Then _
            Call Skynet.Throw(Me, "Public Function IndexOf").InvalidInput("(Count < 0) Or (Index > myClass.Size - Count + myClass.Items.LowerBound)")
    End If
    
    IndexOf = myClass.Items.IndexOf(Value:=Item, StartIndex:=Index, Count:=Count)
    
End Function

Public Sub Reverse( _
                      Optional ByVal Index As Long, _
                      Optional ByVal Count As Long)
    
    If (Index = 0) Then
        Index = myClass.Items.LowerBound
    Else
        If (Index > myClass.Size) Or (Index < myClass.Items.LowerBound) Then _
            Call Skynet.Throw(Me, "Public Sub Reverse").IndexOutofRange("(Index > myClass.Size) Or (Index < myClass.Items.LowerBound)")
    End If
    
    If (Count = 0) Then
        Count = myClass.Size - Index + myClass.Items.LowerBound
    Else
        If (Count < 0) Or (Index > myClass.Size - Count + myClass.Items.LowerBound) Then _
            Call Skynet.Throw(Me, "Public Sub Reverse").InvalidInput("(Count < 0) Or (Index > myClass.Size - Count + myClass.Items.LowerBound)")
    End If
    
    Call myClass.Items.Reverse(Index:=Index, Length:=Count)
    
End Sub

'Sets the capacity of this list to the Size of the list. This method can
'be used to minimize a list's memory overhead once it is known that no
'new elements will be added to the list. To completely clear a list and
'release all memory referenced by the list, execute the following
'statements:
'
'list.Clear();
'list.TrimExcess();
Public Sub TrimExcess()
     
    If myClass.Size < Int(CDbl(myClass.Items.Length) * 0.9) Then _
        Me.Capacity = myClass.Size

End Sub

Private Sub LoadIterator( _
                            ByRef Iterator As Iterator, _
                            ByVal Items As GenericArray, _
                            ByVal StartIndex As Long, _
                            ByVal Count As Long)
 
    With Iterator
        Set .Current = Nothing
        Set .Items = Items
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
    
    Dim c As Member
    
    With c
        Set .Items = Skynet.Clone(myClass.Items)
        .Size = myClass.Size
    End With
    
    Set IGeneric_Clone = Skynet.CreateInstance(New GenericList, VarPtr(c), LenB(c))
   
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
Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericList): End Function
Private Function IGeneric_ToString() As String
    With myClass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Capacity: " & .Items.Length & vbNewLine & _
                            "Elements: " & .Size
    End With
End Function

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_BaseIndex() As Long: IGenericReadOnlyList_BaseIndex = myClass.Items.LowerBound: End Property
Private Property Get IGenericReadOnlyList_Count() As Long: IGenericReadOnlyList_Count = myClass.Size: End Property
Private Function IGenericReadOnlyList_IndexOf(ByVal Item As IGeneric) As Long: IGenericReadOnlyList_IndexOf = myClass.Items.IndexOf(Value:=Item, Count:=myClass.Size): End Function
Private Property Get IGenericReadOnlyList_Item(ByVal Index As Long) As IGeneric
    
    If (Index > myClass.Size) Then _
        Call Skynet.Throw(Me, "Private Property Get IGenericReadOnlyList_Item").IndexOutofRange("Index not valid")

    Set IGenericReadOnlyList_Item = myClass.Items(Index)
    
End Property

Private Sub IGenericReadOnlyList_Clear(): Call Me.Clear: End Sub
Private Sub IGenericReadOnlyList_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call Me.CopyTo(Index:=myClass.Items.LowerBound, GArray:=GArray, ArrayIndex:=ArrayIndex, Elements:=myClass.Size): End Sub
Private Function IGenericReadOnlyList_Iterator() As IGenericIterator: Set IGenericReadOnlyList_Iterator = Me.Iterator: End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = False: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = myClass.IsReadOnly: End Property
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = Me.ToArray: End Function
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.IndexOf(Item:=Item): End Function
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Iterator: End Function
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call Me.CopyTo(Index:=myClass.Items.LowerBound, GArray:=GArray, ArrayIndex:=ArrayIndex, Elements:=myClass.Size): End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myClass.Iterator.Current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    With myClass.Iterator
        If .Index = .EndIndex Then
            If .HasLoaded = False Then _
                Call Skynet.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator
            
            Set Out = Nothing
            Set .Current = Nothing
            Exit Function
        End If
        
        Set .Current = .Items(InterlockedIncrement(.Index)): Set Out = .Current
    End With
    
    IGenericIterator_HasNext = True
    
End Function

Private Sub IGenericIterator_Reset()
    With myClass.Iterator
        If .HasLoaded = False Then _
            Call Skynet.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
                
        Call LoadIterator(myClass.Iterator, .Items, .StartIndex, .Count)
    End With
End Sub

'IPrototype
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericList: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GenericList: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.Size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property

