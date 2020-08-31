VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<T>Enumerable")
Option Compare Database

Option Explicit
'@PredeclaredId

'Implements IPrototype
'Implements IGeneric
Implements IGenericDictionary
Implements IGenericCollection
'Implements IGenericEnumerator

Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Var() As Any) As Long

Private Const MAX_CALLSTACK_SIZE As Long = 1000

Private Enum QueryType
    Keys_
    Values_
    Nodes_
End Enum

Private Type Member
    Root As GenericTreeNode
    count As Long
End Type
Private myClass As Member

Private Type InternalStack
    Size As Long
    Items() As GenericTreeNode
End Type

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get MaxHeight() As Long: MaxHeight = 2 * Skynet.Logn(myClass.count + 2, 2): End Property
Public Property Get count() As Long: count = myClass.count: End Property

Public Property Set Item(ByVal Key As IGeneric, ByVal Value As IGeneric)

    Dim Node As GenericTreeNode: Set Node = FindNode(Key)
    
    If (Node Is Nothing) Then
         Call Me.Add(Key, Value)
    Else
        Set Node.Key = Key
        Set Node.Value = Value
    End If
   
End Property

Public Property Get Item(ByVal Key As IGeneric) As IGeneric
Attribute Item.VB_UserMemId = 0

    Dim Node As GenericTreeNode: Set Node = FindNode(Key)
    
    If (Node Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Property Get Item").NoSuchElement
       
    Set Item = Node.Value

End Property

Public Function GetMax() As IGeneric
    
    Dim Node As GenericTreeNode: Set Node = myClass.Root
    
    If Not (Node Is Nothing) Then
        Do While Not (Node.Right Is Nothing)
            Set Node = Node.Right
        Loop
    End If
    
    Set GetMax = GenericPair(Node.Key, Node.Value)

End Function

Public Function GetMin() As IGeneric
    
    Dim Node As GenericTreeNode: Set Node = myClass.Root
    
    If Not (Node Is Nothing) Then
        Do While Not (Node.Left Is Nothing)
            Set Node = Node.Left
        Loop
    End If
    
    Set GetMin = GenericPair(Node.Key, Node.Value)

End Function

Public Function GetValues(Optional ByVal Reverse As Boolean) As GenericArray
   
    Dim GArray As GenericArray
    Set GArray = GenericArray.Build(myClass.count)
    
    Dim Index As Long: Index = GArray.LowerBound
    
    Dim Stack As InternalStack
    ReDim Stack.Items(Me.MaxHeight)

    Dim Node As GenericTreeNode
    Dim Current As GenericTreeNode
    
    Set Current = myClass.Root
    If Reverse Then
        Do While Not (Current Is Nothing)
            Call Push(Stack, Current)
            Set Current = Current.Right
        Loop
        Do While Not Stack.Size = 0
            Set Current = Pop(Stack)
            Set GArray(Index) = Current.Value
            Index = Index + 1
            Set Node = Current.Left
            Do While Not (Node Is Nothing)
                Call Push(Stack, Node)
                Set Node = Node.Right
            Loop
        Loop
    Else
        Do While Not (Current Is Nothing)
            Call Push(Stack, Current)
            Set Current = Current.Left
        Loop
        Do While Not Stack.Size = 0
            Set Current = Pop(Stack)
            Set GArray(Index) = Current.Value
            Index = Index + 1
            Set Node = Current.Right
            Do While Not (Node Is Nothing)
                Call Push(Stack, Node)
                Set Node = Node.Left
            Loop
        Loop
    End If
    
    Set GetValues = GArray

End Function

Public Function GetKeys(Optional ByVal Reverse As Boolean) As GenericArray
   
    Dim GArray As GenericArray
    Set GArray = GenericArray.Build(myClass.count)
    
    Dim Index As Long: Index = GArray.LowerBound
    
    Dim Stack As InternalStack
    ReDim Stack.Items(Me.MaxHeight)

    Dim Node As GenericTreeNode
    Dim Current As GenericTreeNode
    
    Set Current = myClass.Root
    If Reverse Then
        Do While Not (Current Is Nothing)
            Call Push(Stack, Current)
            Set Current = Current.Right
        Loop
        Do While Not Stack.Size = 0
            Set Current = Pop(Stack)
            Set GArray(Index) = Current.Key
            Index = Index + 1
            Set Node = Current.Left
            Do While Not (Node Is Nothing)
                Call Push(Stack, Node)
                Set Node = Node.Right
            Loop
        Loop
    Else
        Do While Not (Current Is Nothing)
            Call Push(Stack, Current)
            Set Current = Current.Left
        Loop
        Do While Not Stack.Size = 0
            Set Current = Pop(Stack)
            Set GArray(Index) = Current.Key
            Index = Index + 1
            Set Node = Current.Right
            Do While Not (Node Is Nothing)
                Call Push(Stack, Node)
                Set Node = Node.Left
            Loop
        Loop
    End If
    
    Set GetKeys = GArray
    
End Function

Private Sub AddValuesToArray(ByVal Node As GenericTreeNode, ByVal GArray As GenericArray, ByRef Index As Long)

    If Not (Node.Left Is Nothing) Then _
        Call AddValuesToArray(Node.Left, GArray, Index)
    
    Set GArray(Index) = Node.Value: Index = Index + 1
    
    If Not (Node.Right Is Nothing) Then _
        Call AddValuesToArray(Node.Right, GArray, Index)
    
End Sub

Private Sub AddKeysToArray(ByRef Node As GenericTreeNode, ByVal GArray As GenericArray, ByRef Index As Long)
    
    If Not (Node.Left Is Nothing) Then _
        Call AddKeysToArray(Node.Left, GArray, Index)
    
    Set GArray(Index) = Node.Key: Index = Index + 1
    
    If Not (Node.Right Is Nothing) Then _
        Call AddKeysToArray(Node.Right, GArray, Index)
    
End Sub

Private Sub InOrderTreeWalk( _
                                ByVal GArray As GenericArray, _
                                ByVal Index As Long, _
                                ByVal Reverse As Boolean, _
                                ByVal Action As String)
    
    Dim Stack As InternalStack
    ReDim Stack.Items(Me.MaxHeight)

    Dim Node As GenericTreeNode
    Dim Current As GenericTreeNode
    
    Set Current = myClass.Root
    If Reverse Then
        Do While Not (Current Is Nothing)
            Call Push(Stack, Current)
            Set Current = Current.Right
        Loop
        Do While Not Stack.Size = 0
            Set Current = Pop(Stack)
            Call CallByName(Me, Action, VbMethod, Current, GArray, Index)
            Index = Index + 1
            Set Node = Current.Left
            Do While Not (Node Is Nothing)
                Call Push(Stack, Node)
                Set Node = Node.Right
            Loop
        Loop
    Else
        Do While Not (Current Is Nothing)
            Call Push(Stack, Current)
            Set Current = Current.Left
        Loop
        Do While Not Stack.Size = 0
            Set Current = Pop(Stack)
            Call CallByName(Me, Action, VbMethod, Current, GArray, Index)
            Index = Index + 1
            Set Node = Current.Right
            Do While Not (Node Is Nothing)
                Call Push(Stack, Node)
                Set Node = Node.Left
            Loop
        Loop
    End If
End Sub

Public Function Build() As GenericTree

   'Set myClass.Root = GenericTreeNode.Build(Nothing, Nothing)
    Set Build = New GenericTree 'Skynet.Create(GenericTree)

End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
End Sub

Public Sub Clear()
    Set myClass.Root = Nothing
    myClass.count = 0
End Sub

Public Function Contains(ByVal Key As IGeneric) As Boolean
    Contains = Not FindNode(Key) Is Nothing
End Function

Public Function TryGetValue(ByVal Key As IGeneric, ByRef Value As IGeneric) As Boolean

    Dim Node As GenericTreeNode: Set Node = FindNode(Key)
    
    If (Node Is Nothing) Then _
        Exit Function
        
    Set Value = Node.Value
    TryGetValue = True

End Function

Private Function FindNode(ByVal Key As IGeneric) As GenericTreeNode
    
    Dim Current As GenericTreeNode: Set Current = myClass.Root
    Dim Order As Long
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function FindNode").KeyIsNothing
    
    Do While Not (Current Is Nothing)
       Order = Current.Key.CompareTo(Key)
       If (Order = 0) Then _
           Set FindNode = Current: Exit Function
    
        Set Current = IIf(Order < 0, Current.Left, Current.Right)
    Loop
    
End Function

Private Sub ReplaceNode( _
                            ByVal Match As GenericTreeNode, _
                            ByVal ParentOfMatch As GenericTreeNode, _
                            ByVal Succesor As GenericTreeNode, _
                            ByVal ParentOfSuccesor As GenericTreeNode)

    If (Succesor Is Match) Then
        If Not (Match.Right Is Nothing) Then _
            Call Skynet.Throw(Me, "Private Sub ReplaceNode").NullPointer("Right child must be null!")

        Set Succesor = Match.Left
    Else
    
        If Not (Succesor.Right Is Nothing) Then _
            Succesor.Right.IsRed = True
        
        If Not (ParentOfSuccesor Is Match) Then _
            Set ParentOfSuccesor.Left = Succesor.Right: _
            Set Succesor.Right = Match.Right
            
        Set Succesor.Left = Match.Left

    End If
        
    If Not (Succesor Is Nothing) Then _
        Succesor.IsRed = Match.IsRed
                
    Call ReplaceChildOfNodeOrRoot(ParentOfMatch, Match, Succesor)
                   
End Sub

Public Function Add(ByVal Key As IGeneric, ByVal Value As IGeneric) As Boolean
    
    Dim Node As GenericTreeNode: Set Node = GenericTreeNode(Key, Value, False)
  
    Dim Current As GenericTreeNode
    Dim Parent  As GenericTreeNode
    Dim GrandParent As GenericTreeNode
    Dim GreatGrandParent As GenericTreeNode
    Dim Order As Long
    
    If (myClass.Root Is Nothing) Then
        Set myClass.Root = Node
        Call InterlockedIncrement(myClass.count)
        Add = True: Exit Function
    End If
    
    Set Current = myClass.Root
    Do While Not (Current Is Nothing)
        Order = Current.Key.CompareTo(Key)
        If (Order = 0) Then
            myClass.Root.IsRed = False
            Add = False
            Exit Function
        End If
        
        If (Is4Node(Current)) Then
            'Split4Node(Current)
            With Current
                .IsRed = True
                .Left.IsRed = False
                .Right.IsRed = False
            End With
            'IsRed(Parent)
            If Not (Parent Is Nothing) Then _
                If (Parent.IsRed) Then _
                    Call InsertionBalance(Current, Parent, GrandParent, GreatGrandParent)
           
        End If
        
        Set GreatGrandParent = GrandParent
        Set GrandParent = Parent
        Set Parent = Current
        Set Current = IIf(Order > 0, Current.Left, Current.Right)

   Loop
   
'    If (Parent Is Nothing) Then _
'        Call Skynet.Throw(Me, "Private Sub InsertionBalance").NullPointer("Parent cannot be null here!")
'
    If (Order < 0) Then
        Set Parent.Right = Node
    Else
        Set Parent.Left = Node
    End If
    
    If (Parent.IsRed) Then _
        Call InsertionBalance(Node, Parent, GrandParent, GreatGrandParent)
    
    myClass.Root.IsRed = False
    InterlockedIncrement myClass.count
    
    Add = True
            
End Function

Private Sub InsertionBalance( _
                                ByVal Current As GenericTreeNode, _
                                ByRef Parent As GenericTreeNode, _
                                ByVal GrandParent As GenericTreeNode, _
                                ByVal GreatGrandParent As GenericTreeNode)

            
'    If (GrandParent Is Nothing) Then _
'        Call Skynet.Throw(Me, "Private Sub InsertionBalance").NullPointer("Grand parent cannot be null here!")
'
    Dim ParentIsOnRight As Boolean: ParentIsOnRight = GrandParent.Right Is Parent
    Dim CurrentIsOnRight As Boolean: CurrentIsOnRight = Parent.Right Is Current
    Dim NewChildOfGreatGrandParent As GenericTreeNode
    
    If (ParentIsOnRight = CurrentIsOnRight) Then
        If (CurrentIsOnRight) Then
            Set NewChildOfGreatGrandParent = RotateLeft(GrandParent)
        Else
            Set NewChildOfGreatGrandParent = RotateRight(GrandParent)
        End If
    Else
        If (CurrentIsOnRight) Then
            Set NewChildOfGreatGrandParent = RotateLeftRight(GrandParent)
        Else
            Set NewChildOfGreatGrandParent = RotateRightLeft(GrandParent)
        End If
        Set Parent = GreatGrandParent
    End If
    
    GrandParent.IsRed = True
    NewChildOfGreatGrandParent.IsRed = False
    
    Call ReplaceChildOfNodeOrRoot(GreatGrandParent, GrandParent, NewChildOfGreatGrandParent)
    
End Sub

Private Sub ReplaceChildOfNodeOrRoot( _
                                        ByVal Parent As GenericTreeNode, _
                                        ByVal Child As GenericTreeNode, _
                                        ByVal NewChild As GenericTreeNode)

    If Not (Parent Is Nothing) Then
        If (Parent.Left Is Child) Then
            Set Parent.Left = NewChild
        Else
            Set Parent.Right = NewChild
        End If
    Else
        Set myClass.Root = NewChild
    End If

End Sub

Private Function Is4Node(ByVal Node As GenericTreeNode) As Boolean
    If (IsRed(Node.Left)) Then _
        Is4Node = IsRed(Node.Right)
End Function

Private Function IsRed(ByVal Node As GenericTreeNode) As Boolean
    If Not (Node Is Nothing) Then _
        IsRed = Node.IsRed
End Function

Private Function IsBlack(ByVal Node As GenericTreeNode) As Boolean
    If Not (Node Is Nothing) Then _
        IsBlack = Not Node.IsRed
End Function

Private Function IsNothingOrBlack(ByVal Node As GenericTreeNode) As Boolean
    If (Node Is Nothing) Then _
        IsNothingOrBlack = True: Exit Function
    
    IsNothingOrBlack = Not Node.IsRed
End Function

Private Sub Split4Node(ByVal Node As GenericTreeNode)
    
    With Node
        .IsRed = True
        .Left.IsRed = False
        .Right.IsRed = False
    End With
    
End Sub

Private Function RotateLeft(ByVal Node As GenericTreeNode) As GenericTreeNode

    Dim x  As GenericTreeNode
    With Node
        Set x = .Right
        Set .Right = x.Left
    End With
    
    Set x.Left = Node
    Set RotateLeft = x
    
End Function

Private Function RotateLeftRight(ByVal Node As GenericTreeNode) As GenericTreeNode

    Dim Child  As GenericTreeNode
    Dim GrandChild  As GenericTreeNode
    
    Set Child = Node.Left
    Set GrandChild = Child.Right
            
    With GrandChild
        Set Node.Left = .Right
        Set .Right = Node
        Set Child.Right = .Left
        Set .Left = Child
    End With
   
    Set RotateLeftRight = GrandChild
    
End Function

Private Function RotateRight(ByVal Node As GenericTreeNode) As GenericTreeNode

    Dim x  As GenericTreeNode
  
    Set x = Node.Left
    Set Node.Right = x.Right
    Set x.Right = Node
    
    Set RotateRight = x
    
End Function

Private Function RotateRightLeft(ByVal Node As GenericTreeNode) As GenericTreeNode

    Dim Child  As GenericTreeNode
    Dim GrandChild  As GenericTreeNode
    
    Set Child = Node.Right
    Set GrandChild = Child.Left
            
    With GrandChild
        Set Node.Right = .Left
        Set .Left = Node
        Set Child.Left = .Right
        Set .Right = Child
    End With
   
    Set RotateRightLeft = GrandChild
    
End Function

Private Sub Push(ByRef Stack As InternalStack, ByVal Node As GenericTreeNode)
    With Stack
        Set .Items(.Size) = Node
        .Size = .Size + 1
    End With
End Sub

Private Function Pop(ByRef Stack As InternalStack) As GenericTreeNode
    With Stack
        .Size = .Size - 1
        Set Pop = .Items(.Size)
        Set .Items(.Size) = Nothing
    End With
End Function

'IGeneric
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = IGeneric_HashValue = ObjPtr(Me): End Property
Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member
   
    Virtual.Pointer = VarPtr(c)
    Set IGeneric_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function

Private Sub IGeneric_Dispose()
    myClass = EmptyStructure
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal other As IGeneric) As Boolean
    If other Is Nothing Then IGeneric_Equals = False: Exit Function
    IGeneric_Equals = (Me Is other)
End Function

Private Function IGeneric_CompareTo(ByVal other As IGeneric) As ComparerEnum
    Call Skynet.Throw(Me, "IObject_CompareTo").NotSupported
End Function

Private Function IGeneric_IsRelatedTo(ByVal other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf other Is GenericSortedList): End Function
Private Function IGeneric_ToString() As String
    With myClass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Elements: " & .count
    End With
End Function

'IGenericDictionary
Private Property Get IGenericDictionary_Elements() As IGenericCollection: Set IGenericDictionary_Elements = Me: End Property
Private Property Set IGenericDictionary_Item(ByVal Key As IGeneric, ByVal Value As IGeneric): Set Me.Item(Key) = Value: End Property
Private Property Get IGenericDictionary_Item(ByVal Key As IGeneric) As IGeneric: Set IGenericDictionary_Item = Me.Item(Key): End Property
Private Property Get IGenericDictionary_Values() As IGenericCollection: Set IGenericDictionary_Values = Me.GetValues: End Property
Private Property Get IGenericDictionary_Keys() As IGenericCollection: Set IGenericDictionary_Keys = Me.GetKeys: End Property
Private Function IGenericDictionary_ContainsKey(ByVal Key As IGeneric) As Boolean: IGenericDictionary_ContainsKey = Me.Contains(Key): End Function
Private Function IGenericDictionary_TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: IGenericDictionary_TryGetValue = Me.TryGetValue(Key, Value): End Function
Private Sub IGenericDictionary_Remove(ByVal Key As IGeneric):: End Sub
Private Sub IGenericDictionary_Add(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value): End Sub

Private Sub IGenericDictionary_CopyKeysTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long)

End Sub

Private Sub IGenericDictionary_CopyValuesTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long)

End Sub

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.count: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = False: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = False: End Property
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.Contains(Key:=Item): End Function
Private Function IGenericCollection_ToArray() As GenericArray

    Dim GArray As GenericArray
    Set GArray = GenericArray.Build(myClass.count)
   
    Call IGenericCollection_CopyTo(GArray:=GArray, ArrayIndex:=GArray.LowerBound)
    
    Set IGenericCollection_ToArray = GArray

End Function

Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long)

    If (GArray Is Nothing) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericCollection_CopyTo").NullPointer("GArray Is Nothing")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericCollection_CopyTo").IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < Me.count) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericCollection_CopyTo").InvalidInput("(GArray.Length - Index < Me.Count)")
        
    Dim Stack As InternalStack
    ReDim Stack.Items(Me.MaxHeight)

    Dim Node As GenericTreeNode
    Dim Current As GenericTreeNode
    
    Set Current = myClass.Root
    Do While Not (Current Is Nothing)
        Call Push(Stack, Current)
        Set Current = Current.Left
    Loop
    Do While Not Stack.Size = 0
        Set Current = Pop(Stack)
        Set GArray(ArrayIndex) = GenericPair(Current.Key, Current.Value)
        ArrayIndex = ArrayIndex + 1
        Set Node = Current.Right
        Do While Not (Node Is Nothing)
            Call Push(Stack, Node)
            Set Node = Node.Left
        Loop
    Loop
    
End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_NewEnumerator() As IGenericEnumerator

End Function

'IGenericEnumerator
Private Property Get IGenericEnumerator_Current() As IGeneric

End Property

Private Function IGenericEnumerator_MoveNext() As Boolean

End Function

Private Sub IGenericEnumerator_Reset()

End Sub

Private Function IGenericList_NewEnumerator() As IGenericEnumerator

End Function

'IPrototype
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericTree: End Function
Private Property Get IPrototype_Default() As Object: Set IPrototype_Default = GenericTree: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property




