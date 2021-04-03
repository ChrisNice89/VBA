VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Implements IGeneric

Private Type Member
    State As Byte
    Value As IGeneric
    Parent As GenericNode
    Left As GenericNode
    Right As GenericNode
    Count As Long
End Type
Private myclass As Member

Public Property Get IsLeaf() As Boolean
    If myclass.Left Is Nothing Then _
        If myclass.Right Is Nothing Then _
            IsLeaf = True
End Property

Public Property Get State() As Byte: State = myclass.State: End Property
Public Property Let State(ByVal Value As Byte):  myclass.State = Value: End Property

Public Property Get Count() As Long: Count = myclass.Count: End Property
Public Property Let Count(ByVal Value As Long):  myclass.Count = Value: End Property

Public Property Get Left() As GenericNode: Set Left = myclass.Left: End Property
Public Property Set Left(ByVal l As GenericNode): Set myclass.Left = l: End Property

Public Property Get Right() As GenericNode: Set Right = myclass.Right: End Property
Public Property Set Right(ByVal r As GenericNode): Set myclass.Right = r: End Property

Public Property Get Parent() As GenericNode: Set Parent = myclass.Parent: End Property
Public Property Set Parent(ByVal P As GenericNode): Set myclass.Parent = P: End Property

Public Property Get Value() As IGeneric: Set Value = myclass.Value: End Property
Public Property Set Value(ByVal V As IGeneric): Set myclass.Value = V: End Property

Public Sub AddCount(ByVal Count As Long): myclass.Count = myclass.Count + Count: End Sub

Public Function Copy() As GenericNode
    
    Dim newNode As GenericNode
    Set newNode = New GenericNode
    
    With newNode
        Set .Value = myclass.Value
        .State = myclass.State
        .Count = myclass.Count
    End With
    
    Set Copy = newNode
    
End Function

Public Function InOrderNext() As GenericNode
    
    Dim P As GenericNode
    Dim n As GenericNode
    
    If myclass.Right Is Nothing = False Then
        Set P = myclass.Right
        Do While (P.Left Is Nothing) = False
            Set P = P.Left
        Loop
    Else
        Set P = myclass.Parent
        Set n = Me
        Do While (P Is Nothing) = False
            If (P.Right Is n) = False Then _
                Exit Do
            
            Set n = P
            Set P = P.Parent
        Loop
    End If
    
    Set InOrderNext = P
        
End Function

Public Function InOrderPrevious() As GenericNode
    
    Dim P As GenericNode
    Dim n As GenericNode
    
    If myclass.Left Is Nothing = False Then
        Set P = myclass.Left
        Do While (P.Right Is Nothing) = False
            Set P = P.Right
        Loop
    Else
        Set P = myclass.Parent
        Set n = Me
        Do While (P Is Nothing) = False
            If (P.Left Is n) = False Then _
                Exit Do
            
            Set n = P
            Set P = P.Parent
        Loop
    End If
    
    Set InOrderPrevious = P
        
End Function


'Private Sub Class_Terminate()
'    Debug.Print "Node Delete"
'End Sub

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericNode: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GenericNode: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long
    
    Const MAXLONG = 2 ^ 31 - 1
    Dim Hash As Currency
    Hash = 13
    
    If Not myclass.Value Is Nothing Then _
        Hash = (Hash * CDbl(Hash * 27 + myclass.Value.HashValue))
    
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
        .State = myclass.State
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
