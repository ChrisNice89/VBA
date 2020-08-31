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
Implements IPrototype

Private Type Member
    Key As IGeneric
    Value As Object
    Left As GenericTreeNode
    Right As GenericTreeNode
    Parent As GenericTreeNode
    IsRed As Boolean
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get IsRed() As Boolean: IsRed = myClass.IsRed: End Property
Public Property Let IsRed(ByVal Value As Boolean):  myClass.IsRed = Value: End Property

Public Property Get Left() As GenericTreeNode: Set Left = myClass.Left: End Property
Public Property Set Left(ByVal Node As GenericTreeNode): Set myClass.Left = Node: End Property

Public Property Get Right() As GenericTreeNode: Set Right = myClass.Right: End Property
Public Property Set Right(ByVal Node As GenericTreeNode): Set myClass.Right = Node: End Property

Public Property Get Parent() As GenericTreeNode: Set Parent = myClass.Parent: End Property
Public Property Set Parent(ByVal Node As GenericTreeNode): Set myClass.Parent = Node: End Property

Public Property Get Key() As IGeneric: Set Key = myClass.Key: End Property
Public Property Set Key(ByVal k As IGeneric): Set myClass.Key = k: End Property

Public Property Get Value() As Object: Set Value = myClass.Value: End Property
Public Property Set Value(ByVal v As IGeneric): Set myClass.Value = v: End Property

Public Function Build( _
                        ByVal Key As IGeneric, _
                        ByVal Value As Object, _
                        Optional ByVal IsRed As Boolean = False) As GenericTreeNode
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
'    With GenericTreeNode
'        Set myClass.Key = .Key: Set myClass.Value = .Value: myClass.IsRed = .IsRed
'    End With
End Sub

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member: c = myClass

    Virtual.Pointer = VarPtr(c)
    Set IGeneric_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function

Private Sub IGeneric_Dispose(): myClass = EmptyStructure: End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal other As IGeneric) As Boolean
    
    Dim Node As GenericTreeNode: Set Node = other

    If other Is Nothing Then IGeneric_Equals = False: Exit Function
    
    If Not myClass.Value Is Nothing Then _
        If TypeOf myClass.Value Is IGeneric Then _
            IGeneric_Equals = myClass.Key.Equals(Node.Key) And myClass.Value.Equals(Node.Value): Exit Function
 
    IGeneric_Equals = myClass.Key.Equals(Node.Key) And myClass.Value Is Node.Value
    
End Function

Private Function IGeneric_CompareTo(ByVal other As IGeneric) As ComparerEnum
    
    Dim Node As GenericTreeNode: Set Node = other
    
    If other Is Nothing Then _
        IGeneric_CompareTo = IsGreater: Exit Function
    
    IGeneric_CompareTo = myClass.Key.CompareTo(Node.Key)
    
End Function
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = myClass.Key.HashValue: End Property
Private Function IGeneric_IsRelatedTo(ByVal other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf other Is GenericPair): End Function
Private Function IGeneric_ToString() As String
    
    IGeneric_ToString = "Class: " & TypeName$(Me) & vbNewLine & _
                        "[" & _
                            myClass.Key.ToString & _
                            ", " & _
                            myClass.Value.ToString & _
                         "]"
 
End Function

Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericPair: End Function
Private Property Get IPrototype_Default() As Object: Set IPrototype_Default = GenericPair: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property




