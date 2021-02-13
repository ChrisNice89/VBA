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
    Count As Long
End Type
Private myclass As Member

Public Property Get IsLeaf() As Boolean
    If myclass.Left Is Nothing Then _
        If myclass.Right Is Nothing Then _
            IsLeaf = True
End Property
Public Property Get IsRed() As Boolean: IsRed = myclass.IsRed: End Property
Public Property Let IsRed(ByVal Value As Boolean):  myclass.IsRed = Value: End Property

Public Property Get Count() As Long: Count = myclass.Count: End Property
Public Property Let Count(ByVal Value As Long):  myclass.Count = Value: End Property

Public Property Get Left() As GenericNode: Set Left = myclass.Left: End Property
Public Property Set Left(ByVal l As GenericNode): Set myclass.Left = l: End Property

Public Property Get Right() As GenericNode: Set Right = myclass.Right: End Property
Public Property Set Right(ByVal R As GenericNode): Set myclass.Right = R: End Property

Public Property Get Parent() As GenericNode: Set Parent = myclass.Parent: End Property
Public Property Set Parent(ByVal p As GenericNode): Set myclass.Parent = p: End Property

Public Property Get Value() As IGeneric: Set Value = myclass.Value: End Property
Public Property Set Value(ByVal V As IGeneric): Set myclass.Value = V: End Property

Public Sub Invalidate()
    myclass = EmptyStructure
End Sub

Public Function InOrderNext() As GenericNode
    
    Dim p As GenericNode
    Dim N As GenericNode
    
    If myclass.Right Is Nothing = False Then
        Set p = myclass.Right
        Do While (p.Left Is Nothing) = False
            Set p = p.Left
        Loop
    Else
        Set p = myclass.Parent
        Set N = Me
        Do While (p Is Nothing) = False
            If (p.Right Is N) = False Then _
                Exit Do
            
            Set N = p
            Set p = p.Parent
        Loop
    End If
    
    Set InOrderNext = p
        
End Function

Public Function InOrderPrevious() As GenericNode
    
    Dim p As GenericNode
    Dim N As GenericNode
    
    If myclass.Left Is Nothing = False Then
        Set p = myclass.Left
        Do While (p.Right Is Nothing) = False
            Set p = p.Right
        Loop
    Else
        Set p = myclass.Parent
        Set N = Me
        Do While (p Is Nothing) = False
            If (p.Left Is N) = False Then _
                Exit Do
            
            Set N = p
            Set p = p.Parent
        Loop
    End If
    
    Set InOrderPrevious = p
        
End Function

Public Function Build( _
                        ByVal Value As IGeneric, _
                        Optional ByVal Parent As GenericNode, _
                        Optional ByVal IsRed As Boolean = True, _
                        Optional ByVal Count As Long = 1) As GenericNode
Attribute Build.VB_UserMemId = 0
    
    If (Me Is GenericNode) = False Then _
        Call System.Throw(Me, "Public Function Build").BuildExeption
    
    Dim newNode As GenericNode
    Set newNode = New GenericNode
    
    With newNode
        Set .Value = Value
        Set .Parent = Parent
        .IsRed = IsRed
        .Count = Count
    End With
    
    Set Build = newNode
    
End Function

Private Sub Class_Terminate()
'    Debug.Print "Node_Terminate"
End Sub

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericNode: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GenericNode: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
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
    Call Me.Invalidate
End Sub

Private Function EmptyStructure() As Member: End Function
'Todo adjust to DeepCopy
Private Function IGeneric_Clone() As IGeneric

    Dim newNode As GenericNode
    Set newNode = New GenericNode
    
    With newNode
        Set .Value = myclass.Value
        Set .Left = myclass.Left '!
        Set .Right = myclass.Right '!
        Set .Parent = myclass.Parent '!
        .IsRed = myclass.IsRed
        .Count = myclass.Count
    End With
    
    Set IGeneric_Clone = newNode
    
End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

    Dim otherNode As GenericNode
    
    If Other Is Nothing Then _
        Exit Function
        
    Set otherNode = Other
    IGeneric_Equals = myclass.Value.Equals(otherNode.Value)
     
End Function

Private Function IGeneric_ToString() As String
    
    IGeneric_ToString = "Class: " & TypeName$(Me) & " :: " & _
                        "["
    If Not myclass.Value Is Nothing Then
        IGeneric_ToString = IGeneric_ToString & myclass.Value.ToString
    Else
        IGeneric_ToString = IGeneric_ToString & "NULL"
    End If
    
    IGeneric_ToString = IGeneric_ToString & "]"
    
End Function
