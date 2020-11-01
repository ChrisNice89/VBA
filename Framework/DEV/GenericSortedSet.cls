VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSortedSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("<T>Enumerable")

Option Explicit
'@PredeclaredId

Implements IGeneric
Implements IGenericCollection

Private Type Member
    Tree As GenericTree
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Count() As Long: Count = myclass.Tree.Count: End Property
Public Function GetMax() As IGeneric: Set GetMax = myclass.Tree.GetMax: End Function
Public Function GetMin() As IGeneric: Set GetMin = myclass.Tree.GetMin: End Function
Public Function Add(ByVal Value As IGeneric) As Boolean: Add = myclass.Tree.Add(Value, Value): End Function 'Performance Hack
Public Sub AddAll(ByVal Iterator As IGenericIterator)

    Dim P As GenericPair
    Dim Item As IGeneric
    
    With Iterator
        Do While .HasNext(Item)
            If TypeOf Item Is GenericPair Then
                Set P = Item
                Call myclass.Tree.Add(P.Key, P)
            Else
                Call myclass.Tree.Add(Item, Item)
            End If
        Loop
    End With
    
End Sub
Public Function ToArray(Optional ByVal Reverse As Boolean) As GenericArray: Set ToArray = myclass.Tree.GetValues(Reverse:=False): End Function
Public Function Iterator(Optional ByVal Reverse As Boolean) As IGenericIterator: Set Iterator = myclass.Tree.Iterator(t:=Values_, Reverse:=Reverse): End Function
Public Function Remove(ByVal Item As IGeneric) As Boolean: Call myclass.Tree.Remove(Item): End Function
Public Function Contains(ByVal Value As IGeneric) As Boolean: Contains = myclass.Tree.ContainsKey(Value): End Function

Public Function Build() As GenericSortedSet
    
    If (Me Is GenericSortedSet) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption
    
    Set myclass.Tree = GenericTree.Build
    
    Set Build = Skynet.CreateInstance(New GenericSortedSet, Virtual.Pointer, Virtual.Size)

End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
End Sub

Public Sub Clear()
    Set myclass.Tree = Nothing
End Sub

Public Function ElementAt(ByVal Index As Long) As IGeneric
    Set ElementAt = myclass.Tree.ElementAt(Index).Key
End Function

'IGeneric
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GenericSortedSet: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = Virtual.Pointer: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = Virtual.Size: End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericSortedSet):  End Property

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member
 
    Set c.Tree = myclass.Tree
    Set IGeneric_Clone = Skynet.CreateInstance(New GenericSortedSet, VarPtr(c), LenB(c))
    
End Function

Private Sub IGeneric_Dispose()
    myclass = EmptyStructure
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    Call Skynet.Throw(Me, "Private Function IGeneric_Equals").NotImplemented
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    Call Skynet.Throw(Me, "Private Function IGeneric_Equals").NotImplemented
End Function

Private Function IGeneric_ToString() As String
    With myclass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Elements: " & Me.Count
    End With
End Function

Private Sub Class_Terminate()
    Call Skynet.Dispose(Me)
End Sub

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = False: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = False: End Property
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.Contains(Value:=Item): End Function
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = Me.ToArray(Reverse:=False): End Function
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long):: End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Iterator(Reverse:=False): End Function
