VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericTreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<T>Enumerable")
Option Compare Database

Option Explicit
'@PredeclaredId

Implements IGeneric
'Implements GenericPair 'Bug!
Implements IPrototype

Private Type Member
    Key As IGeneric
    Value As IGeneric
    Left As GenericTreeNode
    Right As GenericTreeNode
    IsRed As Boolean
End Type
Private myClass As Member

Private Type Internal
    size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get IsRed() As Boolean: IsRed = myClass.IsRed: End Property
Public Property Let IsRed(ByVal Value As Boolean):  myClass.IsRed = Value: End Property

Public Property Get Left() As GenericTreeNode: Set Left = myClass.Left: End Property
Public Property Set Left(ByVal Node As GenericTreeNode): Set myClass.Left = Node: End Property

Public Property Get Right() As GenericTreeNode: Set Right = myClass.Right: End Property
Public Property Set Right(ByVal Node As GenericTreeNode): Set myClass.Right = Node: End Property

Public Property Get Key() As IGeneric: Set Key = myClass.Key: End Property
Public Property Set Key(ByVal k As IGeneric): Set myClass.Key = k: End Property

Public Property Get Value() As IGeneric: Set Value = myClass.Value: End Property
Public Property Set Value(ByVal v As IGeneric): Set myClass.Value = v: End Property

'GenericPair
Private Function GenericPair_Build(ByVal Key As IGeneric, ByVal Value As Object) As GenericPair: Set GenericPair_Build = GenericPair(Key, Value): End Function
Private Property Get GenericPair_Key() As IGeneric: Set GenericPair_Key = myClass.Key: End Property
Private Property Get GenericPair_Value() As Object: Set GenericPair_Value = myClass.Value: End Property

Public Function Build( _
                        ByVal Key As IGeneric, _
                        ByVal Value As Object, _
                        Optional ByVal IsRed As Boolean = True) As GenericTreeNode
Attribute Build.VB_UserMemId = 0

    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Build").NullPointer("Key is nothing")
    
    Dim n As GenericTreeNode: Set n = New GenericTreeNode
    
    With n
        Set .Key = Key
        Set .Value = Value
        .IsRed = IsRed
    End With
    
    Set Build = n
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Virtual.size = LenB(myClass)
End Sub

Public Function AsPair() As GenericPair: Set AsPair = GenericPair.Build(myClass.Key, myClass.Value): End Function
Private Function IGeneric_Clone() As IGeneric
    
    Dim C As Member: C = myClass 'Hardcopy

    Set IGeneric_Clone = Skynet.CreateInstance(New GenericTreeNode, VarPtr(C), Virtual.size)
    
End Function

Private Sub IGeneric_Dispose(): myClass = EmptyStructure: End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    
    Dim Pair As GenericPair
    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    
    If TypeOf Other Is GenericPair Then
        Set Pair = Other
        If Not myClass.Value Is Nothing Then
            IGeneric_Equals = myClass.Key.Equals(Pair.Key) And myClass.Value.Equals(Pair.Value): Exit Function
        Else
            IGeneric_Equals = myClass.Key.Equals(Pair.Key) And Pair.Value Is Nothing: Exit Function
        End If
    End If
    
    IGeneric_Equals = myClass.Key.Equals(Other)
    
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    
    Dim Pair As GenericPair
    
    If Other Is Nothing Then _
        IGeneric_CompareTo = IsGreater: Exit Function
    
    If TypeOf Other Is GenericPair Then
        Set Pair = Other
        IGeneric_CompareTo = myClass.Key.CompareTo(Pair.Key): Exit Function
    End If
    
    IGeneric_CompareTo = myClass.Key.CompareTo(Other)
    
End Function
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = myClass.Key.HashValue: End Property
Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericPair): End Function
Private Function IGeneric_ToString() As String
    
    IGeneric_ToString = "Class: " & TypeName$(Me) & vbNewLine & _
                        "[" & _
                            myClass.Key.ToString & _
                            ", " & _
                            myClass.Value.ToString & _
                         "]"
 
End Function

Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericPair: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GenericPair: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property
