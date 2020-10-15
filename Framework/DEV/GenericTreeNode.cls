VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericTreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("<T>Enumerable")

Option Explicit
'@PredeclaredId

Implements GenericPair
Implements IGeneric
Implements IPrototype

#If Win64 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst As Any, ByRef src As Any, ByVal length As LongPtr)
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef src As Any, ByVal length As Long)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst As Any, ByRef src As Any, ByVal length As Long)
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef src As Any, ByVal length As Long)
#End If

Private Type Member
    Key As IGeneric
    Value As IGeneric
    Parent As GenericTreeNode
    Left As GenericTreeNode
    Right As GenericTreeNode
    IsRed As Boolean
    Size As Long
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get IsRed() As Boolean: IsRed = myclass.IsRed: End Property
Public Property Let IsRed(ByVal Value As Boolean):  myclass.IsRed = Value: End Property

Public Property Get Size() As Long: Size = myclass.Size: End Property
Public Property Let Size(ByVal Value As Long):  myclass.Size = Value: End Property

Public Property Get Left() As GenericTreeNode: Set Left = myclass.Left: End Property
Public Property Set Left(ByVal L As GenericTreeNode): Set myclass.Left = L: End Property

Public Property Get Right() As GenericTreeNode: Set Right = myclass.Right: End Property
Public Property Set Right(ByVal R As GenericTreeNode): Set myclass.Right = R: End Property

Public Property Get Parent() As GenericTreeNode: Set Parent = myclass.Parent: End Property
Public Property Set Parent(ByVal P As GenericTreeNode): Set myclass.Parent = P: End Property

Public Property Get Key() As IGeneric: Set Key = myclass.Key: End Property
Public Property Set Key(ByVal K As IGeneric): Set myclass.Key = K: End Property

Public Property Get Value() As IGeneric: Set Value = myclass.Value: End Property
Public Property Set Value(ByVal V As IGeneric): Set myclass.Value = V: End Property

Public Function CompareTo(ByVal Other As IGeneric) As ComparerEnum: CompareTo = myclass.Key.CompareTo(Other): End Function

Public Sub Clear()
    With myclass
        Set .Left = Nothing
        Set .Right = Nothing
        Set .Parent = Nothing
    End With
End Sub
'GenericPair
Private Function GenericPair_Build(ByVal Key As IGeneric, ByVal Value As IGeneric) As GenericPair
    If (Me Is GenericTreeNode) = False Then _
        Call Skynet.Throw(Me, "Private Function GenericPair_Build").BuildExeption
End Function
Private Property Get GenericPair_Key() As IGeneric: Set GenericPair_Key = myclass.Key: End Property
Private Property Get GenericPair_Value() As Object: Set GenericPair_Value = myclass.Value: End Property

Public Function Build( _
                        ByVal Key As IGeneric, _
                        ByVal Value As IGeneric, _
                        ByVal Parent As GenericTreeNode, _
                        Optional ByVal IsRed As Boolean = True) As GenericTreeNode

    If (Me Is GenericTreeNode) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption
  
    With myclass
        Set .Key = Key
        Set .Value = Value
        Set .Parent = Parent
        .Size = 1
        .IsRed = IsRed
    End With

'    Dim N As IPrototype
'    Set N = New GenericTreeNode
'    Call CopyMemory(dst:=ByVal N.VirtualPointer, src:=ByVal Virtual.Pointer, Length:=Virtual.Size): Call ZeroMemory(src:=ByVal Virtual.Pointer, Length:=Virtual.Size)
'
'    Set Build = N
   
    Set Build = Skynet.CreateInstance(New GenericTreeNode, Virtual.Pointer, Virtual.Size)
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
End Sub

Private Sub Class_Terminate()

'    Debug.Print "Node_Terminate"
End Sub

'IGeneric
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Sub IGeneric_Dispose()
    
    If myclass.Left Is Nothing = False Then _
        Call Skynet.Dispose(myclass.Left)

    If myclass.Right Is Nothing = False Then _
        Call Skynet.Dispose(myclass.Right)
        
    myclass = EmptyStructure
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member: c = myclass 'Hardcopy

    Set IGeneric_Clone = Skynet.CreateInstance(New GenericTreeNode, VarPtr(c), Virtual.Size)
    
End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

    Dim otherPair As GenericPair
    
    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    
    Set otherPair = Other
    If myclass.Key.Equals(otherPair.Key) Then
        If (myclass.Value Is Nothing) = False Then
            IGeneric_Equals = myclass.Value.Equals(otherPair.Value)
        Else
            IGeneric_Equals = otherPair.Value Is Nothing
        End If
    End If

End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    
    Dim otherPair As GenericPair
    
    If Other Is Nothing Then _
        IGeneric_CompareTo = IsGreater: Exit Function

    Set otherPair = Other
    IGeneric_CompareTo = myclass.Key.CompareTo(otherPair.Key): Exit Function

End Function
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = myclass.Key.HashValue: End Property
Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericPair): End Function
Private Function IGeneric_ToString() As String
    
    IGeneric_ToString = "Class: " & TypeName$(Me) & " :: GenericPair ::" & _
                        "[" & _
                            myclass.Key.ToString & _
                            ", "
                         
                        If (myclass.Value Is Nothing) = False Then
                            IGeneric_ToString = IGeneric_ToString & myclass.Value.ToString & "]"
                        Else
                            IGeneric_ToString = IGeneric_ToString & " NULL]"
                        End If
 
End Function

Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericPair: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GenericPair: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myclass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property
