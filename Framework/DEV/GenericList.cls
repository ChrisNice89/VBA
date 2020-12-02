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
'* Purpose: The GenericList is a collection of IGeneric objects
'*          that can be accessed by index and having methods for sorting, searching, and modifying list.
'*          i.e. GenericList grows dynamically as the elements are added to it.
'*          If the Size of the current elements (including the new element to be added to the GenericList)
'*          is greater than the maximum Size of the array then increase the Size of array.
'*          But the Size of the array can not be increased dynamically.
'*          So, what happens internally is, a new Array is created and
'*          the old array is copied into the new array.
'*          The count,  get, set, and iterator operations run in constant time.
'*          The add operation runs in amortized constant time, that is, adding n elements requires O(n) time.
'*          All of the other operations run in linear time (roughly speaking).
'*          The constant factor is low compared to that for a LinkedList implementation.
'*          This class provides methods to manipulate the size of the array that is used internally to store the list.
'*          (List class is roughly equivalent to a Vector)
'*          Each List instance has a capacity. The capacity is the size of the array used to store the elements in the list.
'*          It is always at least as large as the list size. As elements are added to an List, its capacity grows automatically.
'*          The details of the growth policy are not specified beyond the fact that adding an element has constant amortized time cost.
'*          An application can increase the capacity of a List instance before adding a large number of elements using the
'*          Let Capacity Property. This may reduce the amount of incremental reallocation.
'*
'* Notes:   The important points about GenericList class are:
'*          List is based on GenericArray data structure.
'*          List class can contain duplicate elements.
'*          List class maintains insertion order.
'*          List allows random access because array works at the index basis.
'*          List, manipulation is little bit slower than a LinkedList because a lot of shifting needs
'*          to occur if any element is removed from the list.
'*
'*
'*======================================================================================================================

'@Folder("<T>Enumerable")

Option Explicit
'@PredeclaredId

Implements IGeneric
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
    Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

Private Const S_OK As Long = &H0&
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
    SortOrder As SortOrder
    Iterator As Iterator
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Count() As Long: Count = myclass.Size: End Property
Public Property Get BaseIndex() As Long: BaseIndex = myclass.Items.LowerBound: End Property
Public Property Get SortOrder() As SortOrder: SortOrder = myclass.SortOrder: End Property
Public Property Let SortOrder(ByVal Order As SortOrder):    myclass.SortOrder = Order: End Property
Public Property Get Last() As IGeneric: Set Last = myclass.Items(myclass.Size): End Property
Public Property Set ElementAt(ByVal Index As Long, ByVal Item As IGeneric)

    If (Index > myclass.Size) Then _
        Call Skynet.Throw(Me, "Public Property Get ElementAt").IndexOutofRange("Index not valid")
    
    Set myclass.Items(Index) = Item

End Property

Public Property Get ElementAt(ByVal Index As Long) As IGeneric
Attribute ElementAt.VB_UserMemId = 0

    If (Index > myclass.Size) Then _
        Call Skynet.Throw(Me, "Public Property Get ElementAt").IndexOutofRange("Index not valid")
    
    Set ElementAt = myclass.Items(Index)

End Property

Public Property Get Capacity() As Long: Capacity = myclass.Items.Length: End Property
Public Property Let Capacity(ByVal Value As Long)
    
    If (Value < myclass.Size) Then _
        Call Skynet.Throw(Me, "Public Property Let Capacity").ArgOutOfRange("Value < myClass.Size")
    
    If (myclass.Items Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Property Let Capacity").NullPointer("myClass.Items Is Nothing")
    
    If Not (Value = myclass.Items.Length) Then
        If (Value > 0) Then
            Dim newItems As GenericArray
            Set newItems = GenericArray.Build(Value)
            If (myclass.Size > 0) Then _
                Call CopyMemory(dst:=ByVal newItems.AdressOf(Element:=1), src:=ByVal myclass.Items.AdressOf(Element:=1), Length:=myclass.Items.TotalSize): _
                Call ZeroMemory(dst:=ByVal myclass.Items.AdressOf(Element:=1), Length:=myclass.Items.TotalSize)
           
            Set myclass.Items = newItems
        Else
            Call myclass.Items.Clear
        End If
    End If
End Property

Private Sub EnsureCapacity(ByVal Min As Long)
    
    Dim newCapacity As Long
    
    With myclass.Items
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

Public Function Build(Optional ByVal Capacity As Long = DEFAULT_CAPACITY) As GenericList
    
    If (Me Is GenericList) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption

    If (Capacity < 0) Then _
        Call Skynet.Throw(Me, "Build").ArgOutOfRange("Capacity < 0")

    With myclass
        .SortOrder = ascending
        Set .Items = GenericArray.Build(Capacity)
    End With
    
    Set Build = Skynet.CreateInstance(New GenericList, Virtual.Pointer, Virtual.Size)
    
End Function

Public Function BuildFrom(ByVal Collection As IGenericCollection) As GenericList
    
    If (Me Is GenericList) = False Then _
        Call Skynet.Throw(Me, "Public Function BuildFrom").BuildExeption
    
    If (Collection Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function BuildFrom").NullPointer("(Collection Is Nothing)")
    
    With myclass
        .SortOrder = ascending
        If (Collection Is Nothing) = False Then
            Set .Items = GenericArray.Build(Collection.Count)
            Call Collection.CopyTo(GArray:=.Items, ArrayIndex:=.Items.LowerBound)
            .Size = Collection.Count
        End If
    End With
    
    Set BuildFrom = Skynet.CreateInstance(New GenericList, Virtual.Pointer, Virtual.Size)
    
End Function

Public Function BuildWith(ParamArray GenericValues() As Variant) As GenericList
    
    Dim Items As GenericArray, i As Long, LB As Long: LB = GenericArray.LowerBound
    
    If (Me Is GenericList) = False Then _
        Call Skynet.Throw(Me, "Public Function BuildWith").BuildExeption
    
    If IsMissing(GenericValues) Then _
        Call Skynet.Throw(Me, "Public Function BuildWith").NullPointer("ParamArray GenericValues are missing")
    
    Set Items = GenericArray.Build(UBound(GenericValues) + 1)
    
    For i = LBound(GenericValues) To UBound(GenericValues)
        
        If Not VarType(GenericValues(i)) = vbObject Then _
            Call Skynet.Throw(Me, "Public Function BuildWith").InvalidInput("Datatype {" & TypeName$(GenericValues(i)) & "} not allowed")
        
        If (GenericValues(i) Is Nothing) Then _
            Call Skynet.Throw(Me, "Public Function BuildWith").InvalidInput("Element is nothing")
        
        If TypeOf GenericValues(i) Is IGeneric Then _
            Set Items(i + LB) = GenericValues(i)
      
    Next
    
    With myclass
        .SortOrder = ascending
        Set .Items = Items
        .Size = .Items.Length
    End With
    
    Set BuildWith = Skynet.CreateInstance(New GenericList, Virtual.Pointer, Virtual.Size)
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
End Sub

Private Sub Class_Terminate()
    Call Skynet.Dispose(Me)
End Sub

Public Sub Swap(ByVal i As Long, ByVal j As Long)

    Dim Item As IGeneric
    
    If (i > myclass.Size) Or (j > myclass.Size) Then _
        Call Skynet.Throw(Me, "Public Sub Swap").IndexOutofRange("(i > myClass.Size) Or (j > myClass.Size)")
     
    With myclass.Items
        Set Item = .ElementAt(i)
        Set .ElementAt(i) = .ElementAt(j)
        Set .ElementAt(j) = Item
    End With
    
End Sub

Public Function AsReadOnly() As IGenericReadOnlyList
    Set AsReadOnly = Me 'GenericReadOnlyList.Build(Internal:=myClass.Items, Size:=myClass.Size)
End Function

Public Function Iterator(Optional ByVal Index As Long, Optional ByVal Count As Long) As IGenericIterator
    
    Dim c As Member
    
    If Index = 0 Then _
        Index = myclass.Items.LowerBound
        
    If Count = 0 Then _
        Count = myclass.Size
    
    If (Index < myclass.Items.LowerBound) Or (Index > myclass.Size + myclass.Items.LowerBound) Then _
        Call Skynet.Throw(Me, "Private Sub LoadIterator").IndexOutofRange("(Index < myClass.Items.LowerBound) Or (Index > myClass.Size + myClass.Items.LowerBound)")

    If (Count < 0) Or (Count > myclass.Size - Index + myclass.Items.LowerBound) Then _
        Call Skynet.Throw(Me, "Public Function Iterator").ArgOutOfRange("(Count < 0) Or (Count > myClass.Size - Index + myClass.Items.LowerBound)")
        
    Call LoadIterator(c.Iterator, myclass.Items, Index, Count)
    Set Iterator = Skynet.CreateInstance(New GenericList, VarPtr(c), LenB(c))

End Function

Public Function Add(ByVal Item As IGeneric) As Long

    If (myclass.Size = Me.Capacity) Then _
        Call EnsureCapacity(myclass.Size + 1)
    
    Set myclass.Items(InterlockedIncrement(myclass.Size)) = Item
    Add = myclass.Size

End Function

Public Sub AddAll(ByVal Enumerable As IGenericIterator)
    Call Me.InsertAll(Index:=myclass.Size + myclass.Items.LowerBound, Enumerable:=Enumerable)
End Sub

Public Sub InsertAll(ByVal Index As Long, ByVal Enumerable As IGenericIterator)
    
    Dim Item As IGeneric
    
    If (Enumerable Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Sub InsertAll").NullPointer("Enumerable Is Nothing")
    
    If (Index > myclass.Size + myclass.Items.LowerBound) Then _
        Call Skynet.Throw(Me, "Public Sub InsertAll").IndexOutofRange("Index > myClass.Size")
    
    Index = Index - 1
  
    With Enumerable
        Do While .HasNext(Item): Call Me.Insert(InterlockedIncrement(Index), Item): Loop
        Call .Reset
    End With

End Sub

Public Sub AddCollection(ByVal Collection As IGenericCollection)
    Call Me.InsertCollection(Index:=myclass.Size + myclass.Items.LowerBound, Collection:=Collection)
End Sub

Public Sub InsertCollection(ByVal Index As Long, ByVal Collection As IGenericCollection)
    
    Dim Item As IGeneric
    
    If (Collection Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Sub InsertCollection").NullPointer("Enumerable Is Nothing")
    
    If (Index > myclass.Size + myclass.Items.LowerBound) Then _
        Call Skynet.Throw(Me, "Public Sub InsertCollection").IndexOutofRange("Index > myClass.Size")
    
    If (Collection > 0) Then
        Call EnsureCapacity(myclass.Size + Collection.Count)
    
        With myclass.Items
            If (Index < myclass.Size + .LowerBound) Then _
                Call CopyMemory(dst:=ByVal .AdressOf(Element:=Index + Collection.Count), src:=ByVal .AdressOf(Element:=Index), Length:=.SizeOf(myclass.Size - Index + .LowerBound)): _
                Call ZeroMemory(dst:=ByVal .AdressOf(Element:=Index), Length:=.SizeOf(Collection.Count))
                
        End With
      
        Call Collection.CopyTo(GArray:=myclass.Items, ArrayIndex:=Index)
        myclass.Size = myclass.Size + Collection.Count
    End If
    
End Sub
Public Sub Insert(ByVal Index As Long, ByVal Item As IGeneric)
    
    If (Index > myclass.Size + myclass.Items.LowerBound) Then _
        Call Skynet.Throw(Me, "Public Sub Insert").IndexOutofRange("Index > myClass.Size")
    
    If (myclass.Size = myclass.Items.Length) Then _
        Call EnsureCapacity(myclass.Size + 1)
        
    With myclass.Items
        If (Index < myclass.Size + .LowerBound) Then _
            Call CopyMemory(dst:=ByVal .AdressOf(Element:=Index + 1), src:=ByVal .AdressOf(Element:=Index), Length:=.SizeOf(myclass.Size - Index + .LowerBound)): _
            Call ZeroMemory(dst:=ByVal .AdressOf(Element:=Index), Length:=.SizeOf(1))
         
         Set .ElementAt(Index) = Item: Call InterlockedIncrement(myclass.Size)
    End With

End Sub

Public Function Remove(ByVal Item As IGeneric) As Boolean
    Dim i As Long
    i = Me.IndexOf(Item:=Item)
    
    If (i >= 0) Then _
        Call Me.RemoveAt(i): Remove = True: Exit Function
    
End Function

Public Sub RemoveAt(ByVal Index As Long)

    If (Index > myclass.Size) Then _
        Call Skynet.Throw(Me, "Public Sub RemoveAt").IndexOutofRange("Index > myClass.Size")
    
    InterlockedDecrement myclass.Size
    
    With myclass.Items
        Set .ElementAt(Index) = Nothing
        
        If (Index < myclass.Size + .LowerBound) Then _
            Call CopyMemory(dst:=ByVal .AdressOf(Element:=Index), src:=ByVal .AdressOf(Element:=Index + 1), Length:=.SizeOf(myclass.Size - Index + .LowerBound)): _
            Call ZeroMemory(dst:=ByVal .AdressOf(Element:=myclass.Size), Length:=.SizeOf(1))
            
    End With
    
End Sub

Public Sub RemoveRange(ByVal Index As Long, ByVal Count As Long)

    If (Index > myclass.Size) Then _
        Call Skynet.Throw(Me, "Public Sub RemoveAt").ArgOutOfRange("Index > myClass.Size")
  
End Sub

Public Function GetRange(ByVal Index As Long, ByVal Count As Long) As GenericList
    
    Dim c As Member
     
    With GenericArray

        If (Index < .LowerBound) Or (Index > myclass.Size) Then _
            Call Skynet.Throw(Me, "Public Function GetRange").IndexOutofRange("(Index < .LowerBound) Or (Index > myClass.Size)")

        If (Count < 1) Then _
            Call Skynet.Throw(Me, "Public Function GetRange").ArgOutOfRange("(Count < 1)")
        
        If (myclass.Size < Index + Count) Then _
            Call Skynet.Throw(Me, "Public Function GetRange").InvalidInput("myClass.Size - Index < Count")

        Set c.Items = .Build(Count)
        c.Size = Count
    
        Call .Copy(SrcArray:=myclass.Items, SrcIndex:=Index, DstArray:=c.Items, DstIndex:=c.Items.LowerBound, Elements:=Count)
    End With
    
    Set GetRange = Skynet.CreateInstance(New GenericList, VarPtr(c), Virtual.Size)
   
End Function

Public Sub CopyTo( _
                    ByVal Index As Long, _
                    ByVal GArray As GenericArray, _
                    ByVal ArrayIndex As Long, _
                    Optional ByVal Elements As Long = 0)

    If (Elements = 0) Then _
        Elements = myclass.Size
        
    If (GArray Is Nothing) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericCollection_CopyTo").NullPointer("GArray Is Nothing")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericCollection_CopyTo").IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < Elements) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericCollection_CopyTo").InvalidInput("(GArray.Length - Index < Elements)")
    
    Call GenericArray.Copy(SrcArray:=myclass.Items, SrcIndex:=Index, DstArray:=GArray, DstIndex:=ArrayIndex, Elements:=Elements)

End Sub

'ToArray returns a new Object array containing the contents of the List.
'This requires copying the List, which is an O(n) operation.
Public Function ToArray() As GenericArray

    Dim ga As GenericArray
    Set ga = GenericArray.Build(myclass.Size)

    Call Me.CopyTo(Index:=myclass.Items.LowerBound, GArray:=ga, ArrayIndex:=ga.LowerBound)
    Set ToArray = ga

End Function

Public Function IsEqual(ByVal first As IGenericReadOnlyList, ByVal second As IGenericReadOnlyList) As Boolean

    If first Is Nothing Or second Is Nothing Then _
        Call Skynet.Throw(Me, "Public Function IsEqual").InvalidInput("first Is Nothing Or second Is Nothing")
    
    If first Is second Then _
        IsEqual = True: Exit Function
            
    If Not first.Count = second.Count Then _
        Exit Function
    
    Dim a As IGenericIterator, b As IGenericIterator
    Dim i As IGeneric, j As IGeneric
    
    Set a = first.Iterator: Set b = second.Iterator
    
    Do While (a.HasNext(i) And b.HasNext(j))
'        If Not i Is Nothing Then
            If Not i.Equals(j) Then _
                Exit Function
'        ElseIf Not j Is Nothing Then
'            Exit Function
'        End If
    Loop
    
    IsEqual = True
    
End Function

Public Sub Sort(ByVal Order As SortOrder, Optional ByVal Comparer As IGenericComparer)
    Call myclass.Items.Sort(Order:=Order, Index:=myclass.Items.LowerBound, Length:=myclass.Size, Comparer:=Comparer)
End Sub

Public Sub Clear():
    Call myclass.Items.Clear
    myclass.Size = 0
End Sub

Public Function Contains(Item As IGeneric) As Boolean: Contains = (Me.IndexOf(Item) > myclass.Items.LowerBound): End Function

Public Function BinarySearch( _
                                ByVal Item As IGeneric, _
                                Optional ByVal Comparer As IGenericComparer, _
                                Optional ByVal Index As Long, _
                                Optional ByVal Count As Long) As Long
    
    If (Index = 0) Then
        Index = myclass.Items.LowerBound
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
    
    If (myclass.Size - Index + myclass.Items.LowerBound < Count) Then _
        Call Skynet.Throw(Me, "Public Function BinarySearch").InvalidInput("myClass.Size - Index < Count")
    
    BinarySearch = myclass.Items.BinarySearch(Value:=Item, Index:=Index, Length:=Count, Order:=myclass.SortOrder, Comparer:=Comparer)
    
End Function

Public Function IndexOf( _
                            ByVal Item As IGeneric, _
                            Optional ByVal Index As Long, _
                            Optional ByVal Count As Long) As Long
    
    If (Index = 0) Then
        Index = myclass.Items.LowerBound
    Else
        If (Index > myclass.Size) Or (Index < myclass.Items.LowerBound) Then _
            Call Skynet.Throw(Me, "Public Function IndexOf").IndexOutofRange("(Index > myClass.Size) Or (Index < myClass.Items.LowerBound)")
    End If
    
    If (Count = 0) Then
        Count = myclass.Size - Index + myclass.Items.LowerBound
    Else
        If (Count < 0) Or (Index > myclass.Size - Count + myclass.Items.LowerBound) Then _
            Call Skynet.Throw(Me, "Public Function IndexOf").InvalidInput("(Count < 0) Or (Index > myClass.Size - Count + myClass.Items.LowerBound)")
    End If
    
    IndexOf = myclass.Items.IndexOf(Value:=Item, StartIndex:=Index, Count:=Count)
    
End Function

Public Sub Reverse( _
                      Optional ByVal Index As Long, _
                      Optional ByVal Count As Long)
    
    If (Index = 0) Then
        Index = myclass.Items.LowerBound
    Else
        If (Index > myclass.Size) Or (Index < myclass.Items.LowerBound) Then _
            Call Skynet.Throw(Me, "Public Sub Reverse").IndexOutofRange("(Index > myClass.Size) Or (Index < myClass.Items.LowerBound)")
    End If
    
    If (Count = 0) Then
        Count = myclass.Size - Index + myclass.Items.LowerBound
    Else
        If (Count < 0) Or (Index > myclass.Size - Count + myclass.Items.LowerBound) Then _
            Call Skynet.Throw(Me, "Public Sub Reverse").InvalidInput("(Count < 0) Or (Index > myClass.Size - Count + myClass.Items.LowerBound)")
    End If
    
    Call myclass.Items.Reverse(Index:=Index, Length:=Count)
    
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
     
    If myclass.Size < Int(CDbl(myclass.Items.Length) * 0.9) Then _
        Me.Capacity = myclass.Size

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
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GenericList: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = Virtual.Pointer: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = Virtual.Size: End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericList):  End Property

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member
    
    With c
        Set .Items = Skynet.Clone(myclass.Items)
        .Size = myclass.Size
    End With
    
    Set IGeneric_Clone = Skynet.CreateInstance(New GenericList, VarPtr(c), LenB(c))
   
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
                            "Capacity: " & .Items.Length & vbNewLine & _
                            "Elements: " & .Size
    
    End With
End Function

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_BaseIndex() As Long: IGenericReadOnlyList_BaseIndex = myclass.Items.LowerBound: End Property
Private Property Get IGenericReadOnlyList_Count() As Long: IGenericReadOnlyList_Count = myclass.Size: End Property
Private Function IGenericReadOnlyList_IndexOf(ByVal Item As IGeneric) As Long: IGenericReadOnlyList_IndexOf = myclass.Items.IndexOf(Value:=Item, Count:=myclass.Size): End Function
Private Property Get IGenericReadOnlyList_ElementAt(ByVal Index As Long) As IGeneric
    
    If (Index > myclass.Size) Then _
        Call Skynet.Throw(Me, "Private Property Get IGenericReadOnlyList_ElementAt").IndexOutofRange("Index not valid")

    Set IGenericReadOnlyList_ElementAt = myclass.Items(Index)
    
End Property

Private Sub IGenericReadOnlyList_Clear(): Call Me.Clear: End Sub
Private Sub IGenericReadOnlyList_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call Me.CopyTo(Index:=myclass.Items.LowerBound, GArray:=GArray, ArrayIndex:=ArrayIndex, Elements:=myclass.Size): End Sub
Private Function IGenericReadOnlyList_Iterator() As IGenericIterator: Set IGenericReadOnlyList_Iterator = Me.Iterator: End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = False: End Property
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = Me.ToArray: End Function
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.IndexOf(Item:=Item): End Function
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Iterator: End Function
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call Me.CopyTo(Index:=myclass.Items.LowerBound, GArray:=GArray, ArrayIndex:=ArrayIndex, Elements:=myclass.Size): End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.Current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    With myclass.Iterator
        Do
            If .Index = .EndIndex Then
                If .HasLoaded = False Then _
                    Call Skynet.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator
                
                Set Out = Nothing
                Set .Current = Nothing
                Exit Function
            End If

            Set .Current = .Items.ElementAt(InterlockedIncrement(.Index))
        
        Loop While .Current Is Nothing
        
        Set Out = .Current
        IGenericIterator_HasNext = True
    
    End With
    
End Function

Private Sub IGenericIterator_Reset()
    With myclass.Iterator
        If .HasLoaded = False Then _
            Call Skynet.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
        
        Call LoadIterator(myclass.Iterator, .Items, .StartIndex, .Count)
    End With
End Sub


