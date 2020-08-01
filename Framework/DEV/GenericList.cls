VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<T>Enumerable")
Option Compare Database

Option Explicit
'@PredeclaredId

Implements IPrototype
Implements IGeneric
Implements IGenericList
Implements IGenericCollection
Implements IGenericEnumerator

Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
Private Const DEFAULTCAPACITY As Long = 32

Private Type Member
    Items As GenericArray
    Size As Long
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Last() As IGeneric: Set Last = myClass.Items(myClass.Size): End Property
Public Property Set ElementAt(ByVal Index As Long, ByVal Item As IGeneric): Set myClass.Items(Index) = Item: End Property
Public Property Get ElementAt(ByVal Index As Long) As IGeneric: Set ElementAt = myClass.Items(Index): End Property
Attribute ElementAt.VB_UserMemId = 0
Public Property Get Count() As Long: Count = myClass.Size: End Property
Public Property Get Capacity() As Long: Capacity = myClass.Items.Length: End Property

Public Property Let Capacity(ByVal Value As Long)

    Dim newItems As GenericArray
    
    If (Value < myClass.Size) Then _
        Call Skynet.Push(Me, "Let Capacity").ArgOutOfRange.Raise("Value < myClass.Size")
    
    If (myClass.Items Is Nothing) Then _
        Call Skynet.Push(Me, "Let Capacity").NULLPOINTER.Raise("myClass.Items Is Nothing")
    
    If Not (Value = myClass.Items.Length) Then
        If (Value > 0) Then
            Set newItems = GenericArray.Build(Value)
            
            With myClass.Items
                If (myClass.Size > 0) Then _
                    Call Skynet.CopyData(Dst:=newItems.AdressOf(Element:=1), Src:=.AdressOf(Element:=1), Bytes:=.TotalSize, ClearSource:=True)
            End With
            
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
            newCapacity = IIf(.Length = 0, DEFAULTCAPACITY, .Length * 2)
            
            If (newCapacity > .MaxLength) Then _
                newCapacity = .MaxLength
            
            If (newCapacity < Min) Then _
                newCapacity = Min
        
            Me.Capacity = newCapacity
        End If
    End With
End Sub

Public Function Build( _
                        Optional ByVal Capacity As Long = DEFAULTCAPACITY, _
                        Optional ByVal GArray As GenericArray) As GenericList
  
    If Not (GArray Is Nothing) Then
        With GArray
            Set myClass.Items = GenericArray.Build(.Length)
            Call .CopyTo(GArray:=myClass.Items, ArrayIndex:=.LowerBound, SourceIndex:=.LowerBound)
        End With
        myClass.Size = Count
    Else
        Set myClass.Items = GenericArray.Build(Capacity)
    End If
    
    Set Build = Skynet.Create(Me)
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
End Sub

Private Sub Class_Terminate()
    Call Skynet.Generic(Me).Dispose
End Sub

Public Function AsReadOnly() As GenericReadOnlyList
    Set AsReadOnly = GenericReadOnlyList.Build(Me)
End Function

Public Function Add(ByVal Item As IGeneric) As Long

    If (myClass.Size = Me.Capacity) Then _
        Call EnsureCapacity(myClass.Size + 1)
    
    Set myClass.Items(InterlockedIncrement(myClass.Size)) = Item
    Add = myClass.Size

End Function

Public Function AddAll(ByRef Items() As IGeneric) As GenericList

    
End Function

Public Sub Insert(ByVal Index As Long, ByVal Item As IGeneric)

    If (Index > myClass.Size) Then _
        Call Skynet.Push(Me, "Insert").IndexOutofRange.Raise("Index > myClass.Size")
    
    If (myClass.Size = Me.Capacity) Then _
        Call EnsureCapacity(myClass.Size + 1)
    
    With myClass.Items
      If (Index < .Length) Then _
          Call Skynet.CopyData(Dst:=ByVal .AdressOf(Element:=Index + 1), Src:=ByVal .AdressOf(Element:=Index), Bytes:=.SizeOf(myClass.Size - Index)): _
          Call Skynet.ClearData(Dst:=ByVal .AdressOf(Element:=Index), Bytes:=.SizeOf(1))
    
      Set .ElementAt(Index) = Item
    End With
    Call InterlockedIncrement(myClass.Size)
    
End Sub

Public Sub InsertAll(ByVal Index As Long, ByVal Collection As IGenericCollection)
    
    If (Index <> 0) Then _
        Index = Index - GenericArray.LowerBound
    
    If TypeOf Collection Is GenericArray Then
    
    Else
    
        Dim e As IGenericEnumerator: Set e = Collection.NewEnumerator
        Do While e.MoveNext
            Call Me.Insert(InterlockedIncrement(Index), e.Current)
        Loop
    End If
End Sub

Public Function SubList(ByVal FromIndex As Long, ByVal ToIndex As Long) As GenericList
    
    Dim c As Member
    Dim lb As Long
    
    lb = GenericArray.LowerBound
    
    If (FromIndex < lb) Or (ToIndex < lb) Then _
        Call Skynet.Push(Me, "SubList").IndexOutofRange.Raise("(FromIndex < 0) Or (ToIndex < 0)")

    If (ToIndex < FromIndex) Then _
        Call Skynet.Push(Me, "SubList").IndexOutofRange.Raise("(ToIndex < FromIndex)")
    
    If (myClass.Size - FromIndex + lb < ToIndex) Then _
        Call Skynet.Push(Me, "SubList").InvalidInput.Raise("myClass.Size - FromIndex < ToIndex")
    
    c.Size = ToIndex - FromIndex + lb
    
    Set c.Items = GenericArray.Build(c.Size)
    Call myClass.Items.CopyTo(GArray:=c.Items, ArrayIndex:=lb, SourceIndex:=FromIndex, Elements:=c.Size)

    Virtual.Pointer = VarPtr(c)
    Set SubList = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
    
End Function

Public Function GetRange(ByVal Index As Long, ByVal Count As Long) As GenericList
    
    Dim c As Member
    Dim lb As Long
    
    lb = GenericArray.LowerBound
    
    If (Index < lb) Then _
        Call Skynet.Push(Me, "GetRange").IndexOutofRange.Raise("Index < 0")
    
    If (Count < lb) Then _
        Call Skynet.Push(Me, "GetRange").ArgOutOfRange.Raise("Index < 0")
    
    If (myClass.Size - Index + lb < Count) Then _
        Call Skynet.Push(Me, "GetRange").InvalidInput.Raise("myClass.Size - Index < Count")

    Set c.Items = GenericArray.Build(Count)
    Call myClass.Items.CopyTo(GArray:=c.Items, ArrayIndex:=lb, SourceIndex:=Index, Elements:=Count)
    c.Size = Count
   
    Virtual.Pointer = VarPtr(c)
    Set GetRange = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
    
End Function

'public class SubList extends AbstractList {
'    private final List original;
'    private final int from;
'    private final int to;
'    public SubList(List original, int from, int to) {
'        this.original = original;
'        this.from = from;
'        this.to = to;
'    }
'
'    public Object get(int i) {
'        if (i < 0 || i > to - from) {
'            throw new IllegalArguementException();
'        }
'
'        return original.get(from + i);
'    }
'}
'
'public static List subList(List original, int from, int to) {
'    return new SubList(original, from, to);
'}
Public Sub CopyTo( _
                    ByVal Index As Long, _
                    ByVal GArray As GenericArray, _
                    ByVal ArrayIndex As Long, _
                    Optional ByVal Elements As Long = 0)

    If (Elements = 0) Then _
        Elements = myClass.Size
        
    If (myClass.Size - Index + myClass.Items.LowerBound < Elements) Then _
        Call Skynet.Push(Me, "CopyTo").InvalidInput.Raise("myClass.Size - Index < Elements")
    
    Call myClass.Items.CopyTo(GArray:=GArray, ArrayIndex:=ArrayIndex, SourceIndex:=Index, Elements:=Elements)
    
End Sub

'ToArray returns a new Object array containing the contents of the List.
'This requires copying the List, which is an O(n) operation.
Public Function ToArray() As GenericArray: Set ToArray = Skynet.Generic(myClass.Items).Clone: End Function
Public Sub Clear(): myClass.Items.Clear: End Sub
Public Function Contains(Item As IGeneric) As Boolean: Contains = (Me.IndexOf(Item) > myClass.Items.LowerBound): End Function

Public Function IndexOf( _
                            ByVal Item As IGeneric, _
                            Optional ByVal Index As Long, _
                            Optional ByVal Count As Long) As Long
    
    If (Index = 0) Then _
        Index = myClass.Items.LowerBound
    
    If (Count = 0) Then _
        Count = myClass.Size - Index + myClass.Items.LowerBound
    
    If (Index > myClass.Size) Or (Index < myClass.Items.LowerBound) Then _
        Call Skynet.Push(Me, "IndexOf").IndexOutofRange.Raise("(Index > myClass.Size) Or (Index < myClass.Items.LowerBound)")
        
    If (Count < 0) Or (Index > myClass.Size - Count + myClass.Items.LowerBound) Then _
        Call Skynet.Push(Me, "IndexOf").InvalidInput.Raise("(Count < 0) Or (Index > myClass.Size - Count + myClass.Items.LowerBound)")

    IndexOf = myClass.Items.IndexOf(Value:=Item, StartIndex:=Index, Count:=Count, CheckArguments:=False)
    
End Function

Public Sub RemoveAt(ByVal Index As Long)

End Sub

Public Sub Reverse()

End Sub

Public Sub Shuffle()

End Sub

Public Sub Sort( _
                    Optional ByVal Order As SortOrder = SortOrder.Ascending, _
                    Optional ByVal Method As SortMethod = SortMethod.DepthLimitedQuickSort)

End Sub

Public Function BinarySearch(Object As IGeneric) As Long

End Function

'Sets the capacity of this list to the size of the list. This method can
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

'IGeneric
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = IGeneric_HashValue = ObjPtr(Me): End Property
Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member
    Set c.Items = Skynet.Generic(myClass.Items).Clone
    c.Size = myClass.Size
    
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
Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericList): End Function
Private Function IGeneric_ToString() As String
    With myClass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Capacity: " & .Items.Length & vbNewLine & _
                            "Elements: " & .Size
    End With
End Function

'IGenericList
Private Property Get IGenericList_Elements() As IGenericCollection: Set IGenericList_Elements = Me: End Property
Private Property Set IGenericList_Item(ByVal Index As Long, ByVal Item As IGeneric): Set Me(Index) = Item: End Property
Private Property Get IGenericList_Item(ByVal Index As Long) As IGeneric: Set IGenericList_Item = Me(Index): End Property
Private Function IGenericList_IndexOf(ByVal Item As IGeneric) As Long: IGenericList_IndexOf = Me.IndexOf(Item:=Item): End Function

'IGCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = False: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = False: End Property
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Sub IGenericCollection_CopyTo(ByVal GA As GenericArray, ByVal ArrayIndex As Long): Call Me.CopyTo(Index:=1, GArray:=GA, ArrayIndex:=ArrayIndex, Elements:=Me.Count): End Sub

Private Function IGenericCollection_NewEnumerator() As IGenericEnumerator

End Function

'IGenericEnumerator
Private Property Get IGenericEnumerator_Current() As IGeneric

End Property

Private Function IGenericEnumerator_MoveNext() As Boolean

End Function

Private Sub IGenericEnumerator_Reset()

End Sub

Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericList: End Function
Private Property Get IPrototype_Default() As Object: Set IPrototype_Default = GenericList: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property


