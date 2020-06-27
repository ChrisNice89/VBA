VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<T>Enumerable")
Option Compare Database

Option Explicit
'@PredeclaredId

Private Type Member
    Items() As IObject
    Length as Long
    Count As Long
    Capacity As Long
    DefaultCapacity As Long
End Type
Private myClass As Member
'Private Formats() As String

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Let Capacity(ByVal Value As Long)
    Dim newItems() As IObject
    If (Value <> UBound(myClass.Items, 1)) Then
        If (Value > 0) Then
            ReDim newItems(1 To Value) ' Redim Preserve Myclass.Items( 1 to Value)
            If (myClass.Count > 0) Then
		myClass.Capacity=Value
                'Skynet ArrayCopy(newItems,myclass.Items)
            End If
        Else
            'Skynet FillArray with Value Nothing
        End If
    End If
End Property

Public Property Get Capacity() As Long
    Capacity = myClass.Capacity
End Property
Public Property Get Count() As Long
    Count = myClass.Count
End Property

Private Sub Class_Initialize()
    If Not Me Is TList Then
        With myClass
            .DefaultCapacity = 32
            .Capacity = TList.Capacity
            ReDim .Items(1 To .Capacity)
            .Count = TList.Count
        End With
    End If
End Sub

Public Function Build(Optional ByVal Capacity As Long = 32) As TList
    
    myClass.Capacity = Capacity
    Set Build = New TList

End Function

Public Function Add(ByRef Item As IObject) As TList

    If (myClass.Count = myClass.Capacity) Then Call EnsureCapacity(myClass.Count + 1)
    myClass.Count = myClass.Count + 1
    Set myClass.Items(myClass.Count) = Item
    Set Add = Me
    
End Function

Public Function AddObjects(ByRef Items() As IObject) As TList
	
    'Check Array 
    'Length, Dims eat
    'Take all
    If myClass.Count = 0 Then
        myClass.Items = Items
	myClass.Capacity = UBound(Items,1)'Items.Length
        myClass.Count = myClass.Lengt 
    Else
        'Copy
    End If
    
    Set AddObjects = Me
    
End Function

Public Function AsReadOnly() As TReadOnlyList
    Dim Items() As IObject
    Call Me.CopyTo(Items)
    Set AsReadOnly = TReadOnlyList.Build(Items)
End Function
Public Sub CopyTo(ByRef Items() As IObject)
    'Redim Items(1 to myclass.Count)
    Items = myClass.Items
End Sub

Public Sub Clear()
    Erase myClass.Items
End Sub

Public Function Contains(Item As IObject) As Boolean
    Contains = Not (Me.IndexOf(Item) = 0)
End Function

Public Function IndexOf(Item As IObject) As Long

End Function

Public Sub Remove(ByVal Item As IObject)

End Sub

Public Sub Reverse()

End Sub

Public Sub Shuffle()

End Sub

Public Function ToArray() As IObject()
    Dim Items() As IObject
    Call Me.CopyTo(Items)
    ToArray = Items
End Function

Public Sub Sort( _
                    Optional ByVal Order As SortOrder = SortOrder.Ascending, _
                    Optional ByVal Method As SortMethod = SortMethod.DepthLimitedQuickSort)

End Sub

Public Function BinarySearch(Item As IObject) As Long

End Function
       
Private Sub EnsureCapacity(Min As Long)
    
    Dim newCapacity As Long
    
    If Not (UBound(myClass.Items) < Min) Then _
        Exit Sub
        
    newCapacity = IIf(UBound(myClass.Items) = 0, myClass.DefaultCapacity, UBound(myClass.Items) * 2)
    If (newCapacity > Skynet.MaxArraySize) Then newCapacity = Skynet.MaxArraySize
    If (newCapacity < Min) Then newCapacity = Min
    Me.Capacity = newCapacity
End Sub

'Sets the capacity of this list to the size of the list. This method can
'be used to minimize a list's memory overhead once it is known that no
'new elements will be added to the list. To completely clear a list and
'release all memory referenced by the list, execute the following
'statements:
'
'list.Clear();
'list.TrimExcess();
Public Sub TrimExcess()
    Dim treshold As Long
    treshold = Int(CDbl(UBound(myClass.Items)) * 0.9)
    If myClass.Count < treshold Then Me.Capacity = myClass.Count

End Sub
    
