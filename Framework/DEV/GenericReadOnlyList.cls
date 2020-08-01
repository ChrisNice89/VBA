VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericReadOnlyList"
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
Implements IGenericCollection
'Implements IGenericEnumerator

Private Type Member
    List As IGenericList
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Function Build(ByVal List As IGenericList) As GenericReadOnlyList
    
    If (List Is Nothing) Then _
        Call Skynet.Push(Me, "Build").NULLPOINTER.Raise("List is nothing")
        
    Set myClass.List = List
    Set Build = Skynet.Create(Me)
    
End Function
'Class
Public Property Get ElementAt(ByVal Index As Long) As IGeneric: Set ElementAt = myClass.List(Index): End Property
Attribute ElementAt.VB_UserMemId = 0
Public Property Get Count() As Long: Count = myClass.List.Elements.Count: End Property
Public Sub Clear(): Skynet.Generic(Me).Dispose: End Sub
Public Function Contains(ByVal Item As IGeneric) As Boolean: Contains = Me.IndexOf(Item:=Item) >= 0: End Function
Public Function IndexOf(ByVal Item As IGeneric) As Long: IndexOf = myClass.List.IndexOf(Item:=Item): End Function
Public Sub CopyTo(ByVal GA As GenericArray, ByVal ArrayIndex As Long): Call myClass.List.Elements.CopyTo(GA:=GA, ArrayIndex:=ArrayIndex): End Sub

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
End Sub

Private Sub Class_Terminate()
    Call Skynet.Generic(Me).Dispose
End Sub

'IGeneric
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = ObjPtr(Me): End Property
Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member
    Set c.List = Skynet.Generic(c.List).Clone
        
    Virtual.Pointer = VarPtr(c)
    Set IGeneric_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function
Private Sub IGeneric_Dispose(): myClass = EmptyStructure: End Sub
Private Function EmptyStructure() As Member: End Function
Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean: IGeneric_Equals = Skynet.SequenceCompare(Me, Other) = IsEqual: End Function
Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum: IGeneric_CompareTo = Skynet.SequenceCompare(Me, Other): End Function
Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericReadOnlyList): End Function
Private Function IGeneric_ToString() As String
    With myClass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Elements: " & .List.Elements.Count
    End With
End Function

'IGCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = True: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = True: End Property
Private Sub IGenericCollection_CopyTo(ByVal GA As GenericArray, ByVal ArrayIndex As Long): Call Me.CopyTo(GA, ArrayIndex): End Sub
Private Sub IGenericCollection_Clear(): Me.Clear: End Sub
Private Function IGCollection_Contains(ByVal Item As IGeneric) As Boolean: IGCollection_Contains = Me.Contains(Item): End Function
Private Function IGenericCollection_NewEnumerator() As IGenericEnumerator

End Function
'IPrototype
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericReadOnlyList: End Function
Private Property Get IPrototype_Default() As Object: Set IPrototype_Default = GenericReadOnlyList: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property


