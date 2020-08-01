VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSubList"
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

Private Type Member
    Original As IGenericList
    From As Long
    To As Long
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Set ElementAt(ByVal Index As Long, ByVal Item As IGeneric)


    Set myClass.Original(Index) = Item


End Property

Public Property Get ElementAt(ByVal Index As Long) As IGeneric

    Set ElementAt = myClass.Original(Index)
End Property

Public Property Get Count() As Long: Count = myClass.From + myClass.To: End Property

Public Function Build(ByVal Original As IGenericList, ByVal First As Long, ByVal Last As Long) As GenericList

    Set myClass.Original = Original
    myClass.From = First
    myClass.To = Last

    Set Build = Skynet.Create(Me)

End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
End Sub

Private Sub Class_Terminate()
    Call Skynet.Generic(Me).Dispose
End Sub

'IGeneric
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = IGeneric_HashValue = ObjPtr(Me): End Property
Private Function IGeneric_Clone() As IGeneric

    Dim c As Member
    Set c.Original = Skynet.Generic(myClass.Original).Clone

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
                            "Elements: " & .From + .To
    End With
End Function

'IGenericList
Private Property Get IGenericList_Elements() As IGenericCollection: Set IGenericList_Elements = Me: End Property
Private Property Set IGenericList_Item(ByVal Index As Long, ByVal Item As IGeneric): Set Me(Index) = Item: End Property
Private Property Get IGenericList_Item(ByVal Index As Long) As IGeneric: Set IGenericList_Item = Me(Index): End Property
Private Function IGenericList_IndexOf(ByVal Item As IGeneric) As Long


    IGenericList_IndexOf = myClass.Original.IndexOf(Item:=Item)
End Function

'IGCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = False: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = False: End Property

Private Sub IGenericCollection_Clear()


End Sub

Private Sub IGenericCollection_CopyTo(ByVal GA As GenericArray, ByVal ArrayIndex As Long)

    'Call myClass.Original.Elements.CopyTo(Index:=1, GArray:=GA, ArrayIndex:=ArrayIndex, Elements:=Me.Count): End Sub
End Sub

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

