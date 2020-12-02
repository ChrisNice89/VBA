VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("<T>Enumerable")

Option Explicit
'@PredeclaredId

Implements IGeneric

Private Type Member
    Value As IGeneric
    Parent As GenericNode
    Left As GenericNode
    Right As GenericNode
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

Public Property Get Left() As GenericNode: Set Left = myclass.Left: End Property
Public Property Set Left(ByVal L As GenericNode): Set myclass.Left = L: End Property

Public Property Get Right() As GenericNode: Set Right = myclass.Right: End Property
Public Property Set Right(ByVal R As GenericNode): Set myclass.Right = R: End Property

Public Property Get Parent() As GenericNode: Set Parent = myclass.Parent: End Property
Public Property Set Parent(ByVal P As GenericNode): Set myclass.Parent = P: End Property

Public Property Get Value() As IGeneric: Set Value = myclass.Value: End Property
Public Property Set Value(ByVal V As IGeneric): Set myclass.Value = V: End Property

Public Sub Clear()
    With myclass
        Set .Parent = Nothing
        
        If myclass.Left Is Nothing = False Then _
            Call myclass.Left.Clear
        Set .Left = Nothing

        If myclass.Right Is Nothing = False Then _
            Call myclass.Right.Clear
        Set .Right = Nothing
        
    End With
End Sub

Public Function InOrderNext() As GenericNode
    
    Dim P As GenericNode
    Dim N As GenericNode
    
    If myclass.Right Is Nothing = False Then
        Set P = myclass.Right
        Do While (P.Left Is Nothing) = False
            Set P = P.Left
        Loop
    Else
        Set P = myclass.Parent
        Set N = Me
        Do While (P Is Nothing) = False
            If (P.Right Is N) = False Then _
                Exit Do
            
            Set N = P
            Set P = P.Parent
        Loop
    End If
    
    Set InOrderNext = P
        
End Function

Public Function InOrderPrevious() As GenericNode
    
    Dim P As GenericNode
    Dim N As GenericNode
    
    If myclass.Left Is Nothing = False Then
        Set P = myclass.Left
        Do While (P.Right Is Nothing) = False
            Set P = P.Right
        Loop
    Else
        Set P = myclass.Parent
        Set N = Me
        Do While (P Is Nothing) = False
            If (P.Left Is N) = False Then _
                Exit Do
            
            Set N = P
            Set P = P.Parent
        Loop
    End If
    
    Set InOrderPrevious = P
        
End Function

Public Function Build( _
                        ByVal Value As IGeneric, _
                        Optional ByVal Parent As GenericNode, _
                        Optional ByVal IsRed As Boolean = True, _
                        Optional ByVal Size As Long = 1) As GenericNode
    
    If (Me Is GenericNode) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption
    
    If (Value Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Build").NullPointer("(Value Is Nothing)")
        
    With myclass
        Set .Value = Value
        Set .Parent = Parent
        .Size = Size
        .IsRed = IsRed
    End With
   
    Set Build = New GenericNode: myclass = EmptyStructure
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
    With GenericNode
        Set myclass.Value = .Value
        Set myclass.Parent = .Parent
        myclass.Size = .Size
        myclass.IsRed = .IsRed
    End With
End Sub

Private Sub Class_Terminate()
'    Debug.Print "Node_Terminate"
End Sub

'IGeneric
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GenericNode: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = Virtual.Pointer: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = Virtual.Size: End Property
Private Property Get IGeneric_HashValue() As Long
    
      Const MAXLONG = 2 ^ 31 - 1
    Static Hash As Currency
    
    If Hash = 0 Then _
        Hash = (13 * CDbl(13 * 27 + myclass.Value.HashValue))
    
    IGeneric_HashValue = CLng(Hash - (Fix(Hash / MAXLONG) * MAXLONG))
        
End Property

Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean

    Dim otherNode As GenericNode
    
    If Other Is Nothing Then _
        Exit Sub
        
    Set otherNode = Other
    
    IGeneric_IsRelatedTo = myclass.Value.IsRelatedTo(otherNode.Value)

End Property

Private Sub IGeneric_Dispose()
    Call Me.Clear
End Sub

Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = GenericNode(myclass.Value, myclass.Parent, myclass.IsRed, myclass.Size): End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

    Dim otherNode As GenericNode
    
    If Other Is Nothing Then _
        Exit Function
        
    Set otherNode = Other
    IGeneric_Equals = myclass.Value.Equals(otherNode.Value)
     
End Function

Private Function IGeneric_ToString() As String
    
    IGeneric_ToString = "Class: " & TypeName$(Me) & " :: " & _
                        "[" & myclass.Value.ToString & "]"
 
End Function
