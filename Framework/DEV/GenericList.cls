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
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dst As Any, ByRef Src As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef Dst As Any, ByVal Length As LongPtr, ByVal Fill As Byte)
    Private Declare PtrSafe Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As LongPtr, ByRef Dst As LongPtr, ByVal Src As LongPtr) As Long
    Private Declare PtrSafe Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Length As Long, ByRef Dst As LongPtr, ByVal Src As LongPtr) As Long
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
    Private Const POINTERSIZE As Long = 4
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dst As Any, ByRef Src As Any, ByVal Length As Long)
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal Length As Long)
    Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef Dst As Any, ByVal Length As Long, ByVal Fill As Byte)
    Private Declare Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As Long, ByRef Dst As Long, ByVal Src As Long) As Long
    Private Declare Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Length As Long, ByRef Dst As Long, ByVal Src As Long) As Long
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

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

Public Property Get First() As Long: First = myclass.Items.LowerBound: End Property
Public Property Get Last() As Long
    
    Last = myclass.Items.LowerBound
    
    If myclass.Size > 0 Then _
        Last = myclass.Size - Last - 1
   
End Property

Public Property Get Count() As Long: Count = myclass.Size: End Property
Public Property Get SortOrder() As SortOrder: SortOrder = myclass.SortOrder: End Property
Public Property Let SortOrder(ByVal Order As SortOrder):    myclass.SortOrder = Order: End Property
Public Property Set ElementAt(ByVal Index As Long, ByVal Item As IGeneric)
   
    If (Index >= myclass.Size) Then _
        Call System.Throw(Me, "Public Property Set ElementAt").IndexOutofRange("(Index >= myclass.Size)")
    
    Set myclass.Items.ElementAt(Index) = Item

End Property

Public Property Get ElementAt(ByVal Index As Long) As IGeneric
Attribute ElementAt.VB_UserMemId = 0
    
   If (Index >= myclass.Size) Then _
        Call System.Throw(Me, "Public Property Get ElementAt").IndexOutofRange("(Index >= myclass.Size)")
    
    Set ElementAt = myclass.Items.ElementAt(Index)

End Property

Public Property Get Capacity() As Long: Capacity = myclass.Items.Length: End Property
Public Property Let Capacity(ByVal Value As Long)
    
    If (Value < myclass.Size) Then _
        Call System.Throw(Me, "Public Property Let Capacity").ArgOutOfRange("Value < myClass.Size")
    
    If (myclass.Items Is Nothing) Then _
        Call System.Throw(Me, "Public Property Let Capacity").IsNothing("myClass.Items Is Nothing")
    
    If Not (Value = myclass.Items.Length) Then
        If (Value > 0) Then
            Dim newItems As GenericArray
            
            With GenericArray
                Set newItems = .Build(Value)
                If (myclass.Size > 0) Then _
                    Call CopyMemory(Dst:=ByVal newItems.AdressOf(Index:=.LowerBound), Src:=ByVal myclass.Items.AdressOf(Index:=.LowerBound), Length:=myclass.Items.Size): _
                    Call ZeroMemory(Dst:=ByVal myclass.Items.AdressOf(Index:=.LowerBound), Length:=myclass.Items.Size)
            End With
            
            Set myclass.Items = newItems
        Else
            Call myclass.Items.Clear
        End If
    End If
End Property

Private Sub EnsureCapacity(ByVal min As Long)
    
    Dim newCapacity As Long
    
    With myclass.Items
        If (.Length < min) Then
            newCapacity = IIf(.Length = 0, DEFAULT_CAPACITY, .Length * 2)
            
            If (newCapacity > .MaxLength) Then _
                newCapacity = .MaxLength
            
            If (newCapacity < min) Then _
                newCapacity = min
        
            Me.Capacity = newCapacity
        End If
    End With
End Sub

Public Function Build(Optional ByVal Capacity As Long = DEFAULT_CAPACITY) As GenericList
    
    Dim newList As Member
    
    If (Capacity < 0) Then _
        Call System.Throw(Me, "Build").ArgOutOfRange("Capacity < 0")

    newList.SortOrder = Ascending
    Set newList.Items = GenericArray.Build(Capacity)
  
    Set Build = System.CreateInstance(New GenericList, VarPtr(newList), LenB(newList))
    
End Function

Public Function Create(ParamArray GenericValues() As Variant) As GenericList

    Dim i As Long
    Dim Values As GenericArray
    
    If IsMissing(GenericValues) Then _
        Call System.Throw(Me, "Public Function Create").Instantiation("ParamArray GenericValues are missing")
    
    Set Values = GenericArray.Build(UBound(GenericValues) + 1)
    
    With Values
        For i = LBound(GenericValues) To UBound(GenericValues)
            
            If Not VarType(GenericValues(i)) = vbObject Then _
                Call System.Throw(Me, "Public Function Create").InvalidInput("Datatype {" & TypeName$(GenericValues(i)) & "} not allowed")
            
            If (GenericValues(i) Is Nothing) Then _
                Call System.Throw(Me, "Public Function Create").InvalidInput("Element is nothing")
            
            If Not TypeOf GenericValues(i) Is IGeneric Then _
                Call System.Throw(Me, "Public Function Create").InvalidInput("Not TypeOf GenericValues(i) Is IGeneric")
                
            Set .ElementAt(i) = GenericValues(i)
          
        Next
    End With
    
    Set Create = AsList(Values)
    
End Function

Public Function BuildFrom(ByVal Collection As IGenericCollection) As GenericList
    
    Dim newList As Member
    
    If (Collection Is Nothing) Then _
        Call System.Throw(Me, "Public Function BuildFrom").IsNothing("(Collection Is Nothing)")
    
    If (Collection.Count < 0) Then _
        Call System.Throw(Me, "Build").ArgOutOfRange("Collection.Count < 0")
        
    newList.SortOrder = Ascending
    If (Collection Is Nothing) = False Then
        Set newList.Items = GenericArray.Build(Collection.Count)
        Call Collection.CopyTo(GArray:=newList.Items, ArrayIndex:=newList.Items.LowerBound)
        newList.Size = Collection.Count
    End If
   
    Set BuildFrom = System.CreateInstance(New GenericList, VarPtr(newList), LenB(newList))
    
End Function

Private Function AsList(ByRef Items As GenericArray) As GenericList

    Dim newList As Member
    
    newList.SortOrder = Ascending
    Set newList.Items = Items
    newList.Size = Items.Length
    Set Items = Nothing
    
    Set AsList = System.CreateInstance(New GenericList, VarPtr(newList), LenB(newList))

End Function

Private Sub Class_Terminate()
    Call IGeneric_Dispose
End Sub

Public Sub Swap(ByVal i As Long, ByVal j As Long)

    Dim Item As IGeneric
    
    If (i >= myclass.Size) Or (j >= myclass.Size) Then _
        Call System.Throw(Me, "Public Sub Swap").IndexOutofRange("(i >= myclass.Size) Or (j >= myclass.Size)")
     
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
    
    If Count = 0 Then _
        Count = myclass.Size
    
    If (Index < myclass.Items.LowerBound) Then _
        Call System.Throw(Me, "Private Sub LoadIterator").IndexOutofRange("(Index < myclass.Items.LowerBound)")

    If (Count < 0) Or (Count > myclass.Size - Index ) Then _
        Call System.Throw(Me, "Public Function Iterator").ArgOutOfRange("(Count < 0) Or (Count > myclass.Size - Index )")
        
    Call LoadIterator(c.Iterator, myclass.Items, Index, Count)
    Set Iterator = System.CreateInstance(New GenericList, VarPtr(c), LenB(c))

End Function
'Pops an item from the top of the stack.  If the stack is empty, Pop
'throws an InvalidOperationException.
Public Function Pop() As IGeneric
    
    Dim i As Long: i = Me.Last
    
    With Me
        Set Pop = .ElementAt(i)
        Call .RemoveAt(i)
    End With
    
End Function
'Pushes an item to the top of the stack.
Public Function Push(ByVal Value As IGeneric) As GenericList
    Call Me.Add(Value)
    Set Push = Me
End Function
'Returns the top object on the stack without removing it.  If the stack
'is empty, Peek throws an InvalidOperationException.
Public Function Peek() As IGeneric
    With Me
        Set Peek = .ElementAt(Me.Last)
    End With
End Function

Public Function Add(ByVal Item As IGeneric) As Long
    
    If (myclass.Size = Me.Capacity) Then _
        Call EnsureCapacity(myclass.Size + 1)
    
    Set myclass.Items.ElementAt(myclass.Size) = Item
    myclass.Size = myclass.Size + 1
    Add = myclass.Size
    
End Function

Public Sub AddAll(ByVal Enumerable As IGenericIterator)
    Call Me.InsertAll(Index:=myclass.Size, Enumerable:=Enumerable)
End Sub

Public Sub InsertAll(ByVal Index As Long, ByVal Enumerable As IGenericIterator)
    
    Dim Item As IGeneric
    
    If (Enumerable Is Nothing) Then _
        Call System.Throw(Me, "Public Sub InsertAll").IsNothing("(Enumerable Is Nothing)")
    
    With Enumerable
        Do While .HasNext(Item): Call Me.Insert(Index, Item): Index = Index + 1: Loop
    End With

End Sub

Public Sub AddCollection(ByVal Collection As IGenericCollection)
    Call Me.InsertCollection(Index:=myclass.Size, Collection:=Collection)
End Sub

Public Sub InsertCollection(ByVal Index As Long, ByVal Collection As IGenericCollection)
    
    Dim Item As IGeneric
    Dim Count As Long
    
    If (Collection Is Nothing) Then _
        Call System.Throw(Me, "Public Sub InsertCollection").IsNothing("(Collection Is Nothing)")
    
    If (Index > myclass.Size) Then _
        Call System.Throw(Me, "Public Sub InsertCollection").IndexOutofRange("(Index > myclass.Size)")
    
    Count = Collection.Count
    If (Count > 0) Then
        Call EnsureCapacity(myclass.Size + Count)
    
        With myclass.Items
            If (Index < myclass.Size) Then _
                Call CopyMemory(Dst:=ByVal .AdressOf(Index:=Index + Count), Src:=ByVal .AdressOf(Index:=Index), Length:=.SizeOf(myclass.Size - Index)): _
                Call ZeroMemory(Dst:=ByVal .AdressOf(Index:=Index), Length:=.SizeOf(Count))
                
        End With
      
        Call Collection.CopyTo(GArray:=myclass.Items, ArrayIndex:=Index)
        myclass.Size = myclass.Size + Count
    End If
    
End Sub

Public Sub Insert(ByVal Index As Long, ByVal Item As IGeneric)
    
    If (Index > myclass.Size) Then _
        Call System.Throw(Me, "Public Sub Insert").IndexOutofRange("(Index > myclass.Size)")

    If (myclass.Size = myclass.Items.Length) Then _
        Call EnsureCapacity(myclass.Size + 1)
        
    With myclass.Items
        If (Index < myclass.Size) Then _
            Call CopyMemory(Dst:=ByVal .AdressOf(Index:=Index + 1), Src:=ByVal .AdressOf(Index:=Index), Length:=.SizeOf(myclass.Size - Index)): _
            Call ZeroMemory(Dst:=ByVal .AdressOf(Index:=Index), Length:=.SizeOf(1))
         
         Set .ElementAt(Index) = Item
         myclass.Size = myclass.Size + 1
         
    End With

End Sub

Public Function Remove(ByVal Item As IGeneric) As Boolean
    
    Dim i As Long
    i = Me.IndexOf(Item:=Item)
    
    If (i >= 0) Then _
        Call Me.RemoveAt(i): Remove = True: Exit Function
    
End Function

Public Sub RemoveAt(ByVal Index As Long)

    If (Index < myclass.Items.LowerBound) Or (Index >= myclass.Size) Then _
        Call System.Throw(Me, "Public Sub RemoveAt").IndexOutofRange("(Index < myclass.Items.LowerBound) Or (Index >= myclass.Size)")
        
    With myclass.Items
    
        Set .ElementAt(Index) = Nothing
     
        If (Index < InterlockedDecrement(myclass.Size)) Then _
            Call CopyMemory(Dst:=ByVal .AdressOf(Index:=Index), Src:=ByVal .AdressOf(Index:=Index + 1), Length:=.SizeOf(myclass.Size - Index)): _
            Call ZeroMemory(Dst:=ByVal .AdressOf(Index:=myclass.Size), Length:=.SizeOf(1))
    
    End With

End Sub

Public Sub RemoveRange(ByVal Index As Long, ByVal Count As Long)

    If (Index < myclass.Items.LowerBound) Or (Index >= myclass.Size) Then _
        Call System.Throw(Me, "Public Sub RemoveAt").ArgOutOfRange("Index > myClass.Size")
  
End Sub

Public Function GetRange(ByVal Index As Long, ByVal Count As Long) As GenericList
    
    Dim Items As GenericArray

    With GenericArray
        
        If (Count = 0) Then
            Count = myclass.Size
        ElseIf (Count < 0) Then _
            Call System.Throw(Me, "Public Function GetRange").ArgOutOfRange("(Count < 0)")
        End If
        
        If (myclass.Size - Index < Count) Then _
            Call System.Throw(Me, "Public Function GetRange").InvalidInput("(myclass.Size - Index < Count)")

        Set Items = .Build(Count)

        Call .CopyArrays(SrcArray:=myclass.Items, SrcIndex:=Index, DstArray:=Items, DstIndex:=Items.LowerBound, Elements:=Count)
    End With
    
    Set GetRange = AsList(Items)
   
End Function

Public Sub CopyTo( _
                    ByVal Index As Long, _
                    ByVal GArray As GenericArray, _
                    ByVal ArrayIndex As Long, _
                    Optional ByVal Count As Long = 0)

    If (Count = 0) Then _
        Count = myclass.Size
        
    If (myclass.Size - Index < Count) Then _
        Call System.Throw(Me, "Public Sub CopyTo").ArgOutOfRange("(myclass.Size - Index<count)")
    
    Call GenericArray.CopyArrays(SrcArray:=myclass.Items, SrcIndex:=Index, DstArray:=GArray, DstIndex:=ArrayIndex, Elements:=Count)

End Sub

'ToArray returns a new Object array containing the contents of the List.
'This requires copying the List, which is an O(n) operation.
Public Function ToArray() As GenericArray
   
    Set ToArray = GenericArray.Build(myclass.Size)
    Call Me.CopyTo(Index:=myclass.Items.LowerBound, GArray:=ToArray, ArrayIndex:=ToArray.LowerBound)

End Function

Public Function IsEqual(ByVal First As IGenericReadOnlyList, ByVal second As IGenericReadOnlyList) As Boolean

    If First Is Nothing Or second Is Nothing Then _
        Call System.Throw(Me, "Public Function IsEqual").InvalidInput("first Is Nothing Or second Is Nothing")
    
    If First Is second Then _
        IsEqual = True: Exit Function
            
    If Not First.Count = second.Count Then _
        Exit Function
    
    Dim a As IGenericIterator, B As IGenericIterator
    Dim i As IGeneric, j As IGeneric
    
    Set a = First.Iterator: Set B = second.Iterator
    
    Do While (a.HasNext(i) And B.HasNext(j))

        If Not (i Is Nothing) Then
            If i.Equals(j) = False Then _
                Exit Function
        Else
            If Not (j Is Nothing) Then _
                Exit Function
        End If
    Loop
    
    IsEqual = True
    
End Function

Public Sub Sort(ByVal Order As SortOrder, Optional ByVal Comparer As IGenericComparer, Optional ByVal Index As Long = 0, Optional ByVal Count As Long = 0)

    If Count = 0 Then _
        Count = myclass.Size
        
    If (Index < 0) Then _
        Call System.Throw(Me, "Public Sub Sort").InvalidInput("(Index < 0)")
        
    If (Count < 0) Then _
        Call System.Throw(Me, "Public Sub Sort").InvalidInput("(Count < 0)")
    
    If (myclass.Size - Index < Count) Then _
        Call System.Throw(Me, "Public Sub Sort").ArgOutOfRange("(myclass.Size - Index < Count)")
        
    Call myclass.Items.Sort(Order:=Order, Index:=Index, Length:=Count, Comparer:=Comparer)
    
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
    ElseIf (Index < 0) Then
        Call System.Throw(Me, "Public Function BinarySearch").ArgOutOfRange("Index < 0")
    End If
    
    If (Count = 0) Then
        Count = myclass.Size
    ElseIf (Count < 0) Then
        Call System.Throw(Me, "Public Function BinarySearch").ArgOutOfRange("Count < 0")
    End If
    
    If (myclass.Size - Index < Count) Then _
        Call System.Throw(Me, "Public Function BinarySearch").InvalidInput("(myclass.Size - Index < Count)")
    
    BinarySearch = myclass.Items.BinarySearch(Value:=Item, Order:=myclass.SortOrder, Index:=Index, Length:=Count, Comparer:=Comparer)
    
End Function

Public Function IndexOf( _
                            ByVal Item As IGeneric, _
                            Optional ByVal Index As Long, _
                            Optional ByVal Count As Long, _
                            Optional ByVal Comparer As IGenericComparer) As Long
    
    If (Index = 0) Then
        Index = myclass.Items.LowerBound
    ElseIf (Index < 0) Then
        Call System.Throw(Me, "Public Function IndexOf").ArgOutOfRange("Index < 0")
    End If
    
    If (Count = 0) Then
        Count = myclass.Size
    ElseIf (Count < 0) Then
        Call System.Throw(Me, "Public Function IndexOf").ArgOutOfRange("Count < 0")
    End If
    
    If (myclass.Size - Index < Count) Then _
        Call System.Throw(Me, "Public Function IndexOf").InvalidInput("(myclass.Size - Index < Count)")
        
    IndexOf = myclass.Items.IndexOf(Value:=Item, Index:=Index, Length:=Count, Comparer:=Comparer)
    
End Function

Public Function Reverse( _
                            Optional ByVal Index As Long, _
                            Optional ByVal Count As Long) As GenericList
    
    If (Index = 0) Then
        Index = myclass.Items.LowerBound
    ElseIf (Index < 0) Then
        Call System.Throw(Me, "Public Function IndexOf").ArgOutOfRange("Index < 0")
    End If
    
    If (Count = 0) Then
        Count = myclass.Size
    ElseIf (Count < 0) Then
        Call System.Throw(Me, "Public Function IndexOf").ArgOutOfRange("Count < 0")
    End If
    
    If (myclass.Size - Index < Count) Then _
        Call System.Throw(Me, "Public Function IndexOf").InvalidInput("(myclass.Size - Index < Count)")
    
    Set Reverse = myclass.Items.Reverse(Index:=Index, Length:=Count)
    
End Function

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

    Set Iterator.Items = Items
    Iterator.StartIndex = StartIndex
    Iterator.Count = Count
    Iterator.Index = StartIndex - 1
    Iterator.EndIndex = StartIndex + Count 
    Iterator.HasLoaded = True
    
End Sub

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericList: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GenericList: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericList):  End Property

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member

    Set c.Items = myclass.Items.Copy
    c.Size = myclass.Size
 
    Set IGeneric_Clone = System.CreateInstance(New GenericList, VarPtr(c), LenB(c))
   
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
Private Property Get IGenericReadOnlyList_First() As Long: IGenericReadOnlyList_First = Me.First: End Property
Private Property Get IGenericReadOnlyList_Last() As Long: IGenericReadOnlyList_Last = Me.Last: End Property
Private Property Get IGenericReadOnlyList_Count() As Long: IGenericReadOnlyList_Count = myclass.Size: End Property
Private Function IGenericReadOnlyList_IndexOf(ByVal Item As IGeneric) As Long: IGenericReadOnlyList_IndexOf = myclass.Items.IndexOf(Value:=Item, Length:=myclass.Size): End Function
Private Property Get IGenericReadOnlyList_ElementAt(ByVal Index As Long) As IGeneric
    
    If (Index >= myclass.Size) Then _
        Call System.Throw(Me, "Private Property Get IGenericReadOnlyList_ElementAt").IndexOutofRange("(Index >= myclass.Size)")

    Set IGenericReadOnlyList_ElementAt = myclass.Items.elementAt(Index)
    
End Property

Private Sub IGenericReadOnlyList_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call Me.CopyTo(Index:=myclass.Items.LowerBound, GArray:=GArray, ArrayIndex:=ArrayIndex, Count:=myclass.Size): End Sub
Private Function IGenericReadOnlyList_Iterator() As IGenericIterator: Set IGenericReadOnlyList_Iterator = Me.Iterator: End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = False: End Property
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = Me.ToArray: End Function
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.IndexOf(Item:=Item): End Function
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Iterator: End Function
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call Me.CopyTo(Index:=myclass.Items.LowerBound, GArray:=GArray, ArrayIndex:=ArrayIndex, Count:=myclass.Size): End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.Current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    With myclass.Iterator
    
        If .Index < .EndIndex Then
            Set .Current = .Items.ElementAt(InterlockedIncrement(.Index))
            Set Out= .Current
            IGenericIterator_HasNext=True
            exit function
        end if
        
        If .HasLoaded = False Then 
            Call System.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator

        set .current=nothing
        Set Out= .Current
     
    End With
    
End Function

Private Sub IGenericIterator_Reset()
    With myclass.Iterator
        If .HasLoaded = False Then _
            Call System.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
        
        Call LoadIterator(myclass.Iterator, .Items, .StartIndex, .Count)
    End With
End Sub
