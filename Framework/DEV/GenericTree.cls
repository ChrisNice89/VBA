VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("<T>Enumerable")

Option Explicit
'@PredeclaredId

Implements IGeneric
Implements IGenericDictionary
Implements IGenericCollection
Implements IGenericIterator
Implements IPrototype

Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long

Private Type InternalStack
    Size As Long
    Items() As GenericTreeNode
    Reverse As Boolean
End Type

Private Type Iterator
    Type As DictionaryData
    Current As IGeneric
    tree As GenericTree
    Root As GenericTreeNode
    Stack As InternalStack
    HasLoaded As Boolean
End Type

Private Type Member
    Root As GenericTreeNode
    Count As Long
    Iterator As Iterator
    FindCounter As Long
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get MaxHeight() As Long: MaxHeight = 2 * Skynet.Logn(myclass.Count + 1, Base:=2): End Property
Public Property Get Count() As Long: Count = myclass.Count: End Property
Public Property Get FindCounter() As Long: FindCounter = myclass.FindCounter: End Property


Public Property Set Item(ByVal Key As IGeneric, ByVal Value As IGeneric)

    Dim Node As GenericTreeNode: Set Node = FindNode(Key)
    
    If (Node Is Nothing) Then
         Call Me.Add(Key, Value)
    Else
        Set Node.Value = Value
    End If
   
End Property

Public Property Get Item(ByVal Key As IGeneric) As IGeneric

    Dim Node As GenericTreeNode: Set Node = FindNode(Key)
    
    If (Node Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Property Get Item").NoSuchElement
       
    Set Item = Node.Value

End Property

Public Function Build() As GenericTree
    
    If (Me Is GenericTree) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption
    
    Set Build = New GenericTree 'Skynet.Create(GenericTree)
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
End Sub

Public Function GetMax() As IGeneric
    
    Dim Node As GenericTreeNode: Set Node = myclass.Root
    
    If Not (Node Is Nothing) Then
        Do While Not (Node.Right Is Nothing)
            Set Node = Node.Right
        Loop
    End If
    
    Set GetMax = Node.Key

End Function

Public Function GetMin() As IGeneric
    
    Dim Node As GenericTreeNode: Set Node = myclass.Root
    
    If Not (Node Is Nothing) Then
        Do While Not (Node.Left Is Nothing)
            Set Node = Node.Left
        Loop
    End If
    
    Set GetMin = Node.Key

End Function

Public Sub AddAll(ByVal Map As IGenericDictionary)
    
    Dim p As GenericPair

    With Map.Iterator(t:=Pairs_)
        Do While .HasNext(p): Call Me.Add(p.Key, p.Value): Loop
    End With

End Sub

Public Function SelectSmallest(ByVal K As Long) As IGeneric
    Set SelectSmallest = FindKthSmallest(myclass.Root, K).Key
End Function

Private Function FindKthSmallest(ByVal Node As GenericTreeNode, ByVal Rank As Long) As GenericTreeNode
    
    Dim leftSize As Long
    
    If (Node Is Nothing) Then _
        Set FindKthSmallest = Nothing: Exit Function
    
    If Node.Left Is Nothing = False Then _
        leftSize = Node.Left.Size
    
    If leftSize < Rank Then _
        Set FindKthSmallest = FindKthSmallest(Node.Left, Rank): Exit Function
    
    If leftSize > Rank Then _
        Set FindKthSmallest = FindKthSmallest(Node.Right, Rank - leftSize - 1): Exit Function
    
    Set FindKthSmallest = Node

End Function
'private Node select(Node x, int k) {
'        if (x == null) return null;
'        int t = size(x.left);
'        if (t > k) return select(x.left, k);
'        else if (t < k) return select(x.right, k - t - 1);
'        else return x;
'    }

Public Function GetValues(Optional ByVal Reverse As Boolean) As GenericArray
   
    ReDim Values(GenericArray.LowerBound To myclass.Count) As IGeneric
    Dim Index As Long: Index = GenericArray.LowerBound

    Dim Stack As InternalStack
    Dim Current As GenericTreeNode

    Call LoadStack(Stack, Me, myclass.Root, Reverse)

    Do While InOrderTreeWalk(Current, Stack)
        Set Values(Index) = Current.Value
        Index = Index + 1
    Loop
    
    Set GetValues = GenericArray.BuildFrom(Values)
    
End Function

Public Function GetKeys(Optional ByVal Reverse As Boolean) As GenericArray
   
    ReDim Keys(GenericArray.LowerBound To myclass.Count) As IGeneric
    Dim Index As Long: Index = GenericArray.LowerBound

    Dim Stack As InternalStack
    Dim Current As GenericTreeNode

    Call LoadStack(Stack, Me, myclass.Root, Reverse)

    Do While InOrderTreeWalk(Current, Stack)
        Set Keys(Index) = Current.Key
        Index = Index + 1
    Loop
    
    Set GetKeys = GenericArray.BuildFrom(Keys)
    
End Function
'recursive solution is limited in depth and quickly leads to a stack overflow
Private Sub AddValuesToArray(ByVal Node As GenericTreeNode, ByVal GArray As GenericArray, ByRef Index As Long)

    If Not (Node.Left Is Nothing) Then _
        Call AddValuesToArray(Node.Left, GArray, Index)
    
    Set GArray(Index) = Node.Value: Index = Index + 1
    
    If Not (Node.Right Is Nothing) Then _
        Call AddValuesToArray(Node.Right, GArray, Index)
    
End Sub
'recursive solution is limited in depth and quickly leads to a stack overflow
Private Sub AddKeysToArray(ByRef Node As GenericTreeNode, ByVal GArray As GenericArray, ByRef Index As Long)
    
    If Not (Node.Left Is Nothing) Then _
        Call AddKeysToArray(Node.Left, GArray, Index)
    
    Set GArray(Index) = Node.Key: Index = Index + 1
    
    If Not (Node.Right Is Nothing) Then _
        Call AddKeysToArray(Node.Right, GArray, Index)
    
End Sub

'Private Sub InOrderTreeWalk2( _
'                                ByVal GArray As GenericArray, _
'                                ByVal Index As Long, _
'                                ByVal Reverse As Boolean)
'
'    Dim Stack As InternalStack
'    ReDim Stack.Items(Me.MaxHeight)
'    Dim Node As GenericTreeNode
'    Dim Current As GenericTreeNode
'    Set Node = myClass.Root
'
'    If Reverse Then
'        Do While Not (Node Is Nothing)
'            Call Push(Stack, Node)
'            Set Node = Node.Right
'        Loop
'    Else
'        Do While Not (Node Is Nothing)
'            Call Push(Stack, Node)
'            Set Node = Node.Left
'        Loop
'    End If
'
'    If Reverse Then
'        Do While Not Stack.Size = 0
'            Set Current = Pop(Stack)
'            Set GArray(Index) = Current.CopyToPair
'            Index = Index + 1
'            Set Node = Current.Left
'            Do While Not (Node Is Nothing)
'                Call Push(Stack, Node)
'                Set Node = Node.Right
'            Loop
'        Loop
'    Else
'        Do While Not Stack.Size = 0
'            Set Current = Pop(Stack)
'            Set GArray(Index) = Current.CopyToPair
'            Index = Index + 1
'            Set Node = Current.Right
'            Do While Not (Node Is Nothing)
'                Call Push(Stack, Node)
'                Set Node = Node.Left
'            Loop
'        Loop
'    End If
'End Sub

Public Function Iterator(ByVal t As DictionaryData, Optional ByVal Reverse As Boolean) As IGenericIterator
    
    Dim c As Member

    Set c.Root = myclass.Root
    c.Count = myclass.Count
    
    Call LoadIterator(c.Iterator, t, Me, myclass.Root, Reverse)
    Set Iterator = Skynet.CreateInstance(New GenericTree, VarPtr(c), LenB(c))
    
End Function

Public Sub Clear()
    With myclass
        Set .Root = Nothing
        .Count = 0
    End With
End Sub

Public Function Remove(ByVal Key As IGeneric) As Boolean


End Function

Public Function ContainsKey(ByVal Key As IGeneric) As Boolean: ContainsKey = Not FindNode(Key) Is Nothing: End Function
Public Function ContainsValue(ByVal Value As IGeneric) As Boolean
    
    Dim Stack As InternalStack
    Dim Current As GenericTreeNode

    Call LoadStack(Stack, Me, myclass.Root, False)
    
    If Not (Value Is Nothing) Then
        If Value.Equals(Stack.Items(Stack.Size).Value) Then _
            ContainsValue = True: Exit Function
            
        Do While InOrderTreeWalk(Current, Stack)
            If (Value.Equals(Current.Value)) Then _
                ContainsValue = True: Exit Function
        Loop
    Else
        If Stack.Items(Stack.Size).Value Is Nothing Then _
            ContainsValue = True: Exit Function
            
        Do While InOrderTreeWalk(Current, Stack)
            If (Current.Value Is Nothing) Then _
                ContainsValue = True: Exit Function
        Loop
    End If
    
End Function
Public Function TryGetValue(ByVal Key As IGeneric, ByRef Value As IGeneric) As Boolean

    Dim Node As GenericTreeNode: Set Node = FindNode(Key)
    
    If (Node Is Nothing) Then _
        Exit Function
        
    Set Value = Node.Value
    TryGetValue = True

End Function

Private Function FindNode(ByVal Key As IGeneric) As GenericTreeNode
    
    Dim Current As GenericTreeNode: Set Current = myclass.Root
    Dim Order As ComparerEnum
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function FindNode").KeyIsNothing
    
    Do While (Current Is Nothing) = False
        Order = Current.Key.CompareTo(Key)
        myclass.FindCounter = myclass.FindCounter + 1
        If (Order = IsEqual) Then _
            Set FindNode = Current: Exit Function
        
        If Order = IsLower Then
            Set Current = Current.Right
        Else
            Set Current = Current.Left
        End If
        
    Loop
    
End Function

Public Function FindNext(ByVal Key As IGeneric) As IGeneric
    
    Dim Node As GenericTreeNode
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function FindNode").KeyIsNothing
    
    Set Node = myclass.Root
    
    If (Node.Right Is Nothing) = False Then
        Do While (Node.Left Is Nothing) = False
            Set Node = Node.Left
        Loop
    Else
        
        
    End If
    
    Set FindNext = Node
    
End Function

Public Function FindPrevious(ByVal Key As IGeneric) As IGeneric

End Function

Private Function IndexOf(ByVal Key As IGeneric) As Long
    
    Dim Node As GenericTreeNode: Set Node = myclass.Root
    
    Dim Count As Long: Count = 1
    Dim Order As Long
    
    Do While Not (Node Is Nothing)
        Order = Node.Key.CompareTo(Key)
        
        If (Order = 0) Then _
           IndexOf = Count: Exit Function
      
        If (Order > 0) Then
            Set Node = Node.Left
            Count = 2 * Count
        Else
            Set Node = Node.Right
            Count = 2 * Count + 1
        End If
      
    Loop
    
    IndexOf = -1
    
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
    
    Dim Node As GenericTreeNode
    Dim Current As GenericTreeNode
    Dim Parent  As GenericTreeNode
    Dim GrandParent As GenericTreeNode
    Dim GreatGrandParent As GenericTreeNode
    Dim Order As Long, i As Long
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Add").KeyIsNothing
    
    If (myclass.Root Is Nothing) Then _
        Set myclass.Root = GenericTreeNode.Build(Key:=Key, Value:=Value, Size:=0, IsRed:=False): _
        Call InterlockedIncrement(myclass.Count): _
        Add = True: _
        Exit Function
    
    Set Current = myclass.Root
    Do
        Order = Current.Key.CompareTo(Key)
        
        If (Order = 0) Then _
            myclass.Root.IsRed = False: _
            Add = False: _
            Exit Function
        
        If Split4Node(Current) Then
            If (Parent Is Nothing) = False Then
                If Parent.IsRed Then
                    Call InsertionBalance(Current, Parent, GrandParent, GreatGrandParent)
                 
                End If
             End If
        End If
            
        Current.Size = 1 + SizeOf(Current.Left) + SizeOf(Current.Right)
        
        Set GreatGrandParent = GrandParent
        Set GrandParent = Parent
        Set Parent = Current
        
        If Order > 0 Then
            Set Current = Current.Left
        Else
            Set Current = Current.Right
        End If
        
   Loop While (Current Is Nothing) = False
   
'    If (Parent Is Nothing) Then _
'        Call Skynet.Throw(Me, "Private Sub InsertionBalance").NullPointer("Parent cannot be null here!")

    Set Node = GenericTreeNode.Build(Key:=Key, Value:=Value, Size:=0, IsRed:=True)
    
    If (Order < 0) Then
        Set Parent.Right = Node
    Else
        Set Parent.Left = Node
    End If
    
    If (Parent.IsRed) Then _
        Call InsertionBalance(Node, Parent, GrandParent, GreatGrandParent)
    
    myclass.Root.IsRed = False
    myclass.Count = myclass.Count + 1
    
    Add = True
     
End Function

Private Sub InsertionBalance( _
                                ByVal Current As GenericTreeNode, _
                                ByRef Parent As GenericTreeNode, _
                                ByVal GrandParent As GenericTreeNode, _
                                ByVal GreatGrandParent As GenericTreeNode)

'    If (GrandParent Is Nothing) Then _
'        Call Skynet.Throw(Me, "Private Sub InsertionBalance").NullPointer("Grand parent cannot be null here!")

    'Dim ParentIsOnRight As Boolean: ParentIsOnRight = GrandParent.Right Is Parent
    Dim CurrentIsOnRight As Boolean: CurrentIsOnRight = Parent.Right Is Current
    Dim NewChildOfGreatGrandParent As GenericTreeNode
    
    If ((GrandParent.Right Is Parent) = CurrentIsOnRight) Then
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
    
    'Call ReplaceChildOfNodeOrRoot(GreatGrandParent, GrandParent, NewChildOfGreatGrandParent)
    If GreatGrandParent Is Nothing Then
        Set myclass.Root = NewChildOfGreatGrandParent
    Else
        If (GreatGrandParent.Left Is GrandParent) Then
            Set GreatGrandParent.Left = NewChildOfGreatGrandParent
        Else
            Set GreatGrandParent.Right = NewChildOfGreatGrandParent
        End If
    End If
End Sub

Private Sub ReplaceChildOfNodeOrRoot( _
                                        ByVal Parent As GenericTreeNode, _
                                        ByVal Child As GenericTreeNode, _
                                        ByVal NewChild As GenericTreeNode)
    
    If Parent Is Nothing Then _
        Set myclass.Root = NewChild: Exit Sub
    
    If (Parent.Left Is Child) Then
        Set Parent.Left = NewChild
    Else
        Set Parent.Right = NewChild
    End If
        
End Sub

Private Function SizeOf(ByVal Node As GenericTreeNode) As Long
    If Node Is Nothing = False Then _
        SizeOf = Node.Size
End Function
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

Private Function Split4Node(ByVal Node As GenericTreeNode) As Boolean
    With Node
        If (IsRed(.Left)) Then
            If IsRed(.Right) Then
                .IsRed = True
                .Left.IsRed = False
                .Right.IsRed = False
                Split4Node = True
            End If
        End If
    End With
End Function

Private Function RotateLeft(ByVal x As GenericTreeNode) As GenericTreeNode

    Dim y  As GenericTreeNode: Set y = x.Right
   
    Set x.Right = y.Left: Set y.Left = x
    
    y.Size = x.Size
    x.Size = 1 + SizeOf(x.Left) + SizeOf(x.Right)
    
    Set RotateLeft = y
    
End Function

Private Function RotateLeftRight(ByVal Node As GenericTreeNode) As GenericTreeNode

    Dim Child  As GenericTreeNode: Set Child = Node.Left
    Dim GrandChild  As GenericTreeNode: Set GrandChild = Child.Right
    
    With GrandChild
        Set Node.Left = .Right: Set .Right = Node
        Set Child.Right = .Left: Set .Left = Child
        Child.Size = Node.Size
        Node.Size = 1 + SizeOf(Node.Left) + SizeOf(Node.Right)
    End With
   
    Set RotateLeftRight = GrandChild
    
End Function

Private Function RotateRight(ByVal x As GenericTreeNode) As GenericTreeNode

    Dim y  As GenericTreeNode: Set y = x.Left
    
    Set x.Left = y.Right: Set y.Right = x
    
    y.Size = x.Size
    x.Size = 1 + SizeOf(x.Left) + SizeOf(x.Right)
    
    Set RotateRight = y
    
End Function

Private Function RotateRightLeft(ByVal Node As GenericTreeNode) As GenericTreeNode

    Dim Child  As GenericTreeNode: Set Child = Node.Right
    Dim GrandChild  As GenericTreeNode: Set GrandChild = Child.Left
    
    With GrandChild
        Set Node.Right = .Left: Set .Left = Node
        Set Child.Left = .Right: Set .Right = Child
        Child.Size = Node.Size
        Node.Size = 1 + SizeOf(Node.Left) + SizeOf(Node.Right)
    End With
   
    Set RotateRightLeft = GrandChild
    
End Function

Private Function InOrderTreeWalk( _
                                    ByRef Current As GenericTreeNode, _
                                    ByRef Stack As InternalStack) As Boolean
    
    Dim Node As GenericTreeNode

    If Stack.Size = 0 Then _
        Set Current = Nothing: Exit Function
    
    With Stack
        Set Current = .Items(.Size)
        Set .Items(.Size) = Nothing
        .Size = .Size - 1
        If .Reverse = False Then
            Set Node = Current.Right
    
            Do While Not (Node Is Nothing)
                Set .Items(InterlockedIncrement(.Size)) = Node
                Set Node = Node.Left
            Loop
        Else
            Set Node = Current.Left
    
            Do While Not (Node Is Nothing)
                Set .Items(InterlockedIncrement(.Size)) = Node
                Set Node = Node.Right
            Loop
        End If
    End With
    
    InOrderTreeWalk = True

End Function

Private Sub LoadStack(ByRef Stack As InternalStack, ByVal tree As GenericTree, ByVal Node As GenericTreeNode, ByVal Reverse As Boolean)
    
    With Stack
        .Reverse = Reverse
        .Size = 0
        ReDim .Items(1 To tree.MaxHeight)
        
        If .Reverse = False Then
            Do While Not (Node Is Nothing)
                Set .Items(InterlockedIncrement(.Size)) = Node
                Set Node = Node.Left
            Loop
        Else
            Do While Not (Node Is Nothing)
                Set .Items(InterlockedIncrement(.Size)) = Node
                Set Node = Node.Right
            Loop
        End If
    End With
    
End Sub

Private Sub LoadIterator(ByRef e As Iterator, ByVal t As DictionaryData, ByVal tree As GenericTree, ByVal Root As GenericTreeNode, ByVal Reverse As Boolean)
    
    If Root Is Nothing Then _
        Call Skynet.Throw(Me, "Private Sub LoadIterator").NullPointer("Root Is Nothing")
    
    With e
        .Type = t
        Set .Root = Root
        Set .tree = tree
        Call LoadStack(.Stack, Me, .Root, Reverse)
        Set .Current = Nothing
        .HasLoaded = True
    End With
    
End Sub

Private Function CopyOf(ByVal t As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long, ByVal Caller As String) As GenericArray
    
    If (GArray Is Nothing) Then _
        Call Skynet.Throw(Me, Caller).NullPointer("GArray Is Nothing")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call Skynet.Throw(Me, Caller).IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < myclass.Count) Then _
        Call Skynet.Throw(Me, Caller).InvalidInput("(GArray.Length - Index < myClass.Count)")
    
    Dim Stack As InternalStack
    Dim Node As GenericTreeNode
    Dim Current As IGeneric
    Call LoadStack(Stack, Me, myclass.Root, False)

    Select Case True
        Case t = Keys_
            Do While InOrderTreeWalk(Node, Stack)
                Set GArray(ArrayIndex) = Node.Key
                ArrayIndex = ArrayIndex + 1
            Loop
        Case t = Values_
            Do While InOrderTreeWalk(Node, Stack)
                Set GArray(ArrayIndex) = Node.Value
                ArrayIndex = ArrayIndex + 1
            Loop
        Case Else
            Do While InOrderTreeWalk(Node, Stack)
                Set GArray(ArrayIndex) = Node.CopyToPair
                ArrayIndex = ArrayIndex + 1
            Loop
    End Select

    Set CopyOf = GArray
    
End Function

'IGeneric
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = IGeneric_HashValue = ObjPtr(Me): End Property
Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member
    
    With c
        Set .Root = myclass.Root
        .Count = myclass.Count
    End With
    
    Set IGeneric_Clone = Skynet.CreateInstance(New GenericTree, VarPtr(c), LenB(c))
    
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

Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericSortedList): End Function
Private Function IGeneric_ToString() As String
    With myclass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Elements: " & .Count
    End With
End Function

Private Sub Class_Terminate()
    Call Skynet.Dispose(Me)
End Sub

'IGenericDictionary
Private Property Set IGenericDictionary_Item(ByVal Key As IGeneric, ByVal Value As IGeneric): Set Me.Item(Key) = Value: End Property
Private Property Get IGenericDictionary_Item(ByVal Key As IGeneric) As IGeneric: Set IGenericDictionary_Item = Me.Item(Key): End Property
Private Property Get IGenericDictionary_Count() As Long: IGenericDictionary_Count = myclass.Count: End Property
Private Function IGenericDictionary_ContainsKey(ByVal Key As IGeneric) As Boolean: IGenericDictionary_ContainsKey = Me.ContainsKey(Key:=Key): End Function
Private Function IGenericDictionary_ContainsValue(ByVal Value As IGeneric) As Boolean: IGenericDictionary_ContainsValue = Me.ContainsValue(Value:=Value): End Function
Private Function IGenericDictionary_TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: IGenericDictionary_TryGetValue = Me.TryGetValue(Key, Value): End Function
Private Sub IGenericDictionary_Remove(ByVal Key As IGeneric): Call Me.Remove(Key:=Key): End Sub
Private Sub IGenericDictionary_Add(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value): End Sub
Private Sub IGenericDictionary_CopyOf(ByVal t As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(t:=t, GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="Private Sub IGenericDictionary_CopyOf"): End Sub
Private Sub IGenericDictionary_Clear(): Call Me.Clear: End Sub
Private Function IGenericDictionary_Iterator(ByVal t As DictionaryData) As IGenericIterator: Set IGenericDictionary_Iterator = Me.Iterator(t:=t): End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = False: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = False: End Property
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.ContainsKey(Key:=Item): End Function
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = CopyOf(t:=Pairs_, GArray:=GenericArray.Build(myclass.Count), ArrayIndex:=GenericArray.LowerBound, Caller:="Private Function IGenericCollection_ToArray() As GenericArray"): End Function
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(t:=Pairs_, GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="Private Sub IGenericCollection_CopyTo"): End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Iterator(t:=Pairs_, Reverse:=False): End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.Current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    Dim Node As GenericTreeNode
    
    With myclass.Iterator
        If InOrderTreeWalk(Node, .Stack) Then
            
            IGenericIterator_HasNext = True
            
            If .Type = Keys_ Then
                Set .Current = Node.Key
            
            ElseIf .Type = Values_ Then
                Set .Current = Node.Value
            
            Else
                Set .Current = Node.CopyToPair
            End If
            
            Set Out = .Current
            Exit Function
        End If
        
        If .HasLoaded = False Then _
            Call Skynet.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator
        
        Set Out = Nothing
        Set .Current = Nothing
      
    End With
 
End Function

Private Sub IGenericIterator_Reset()
    With myclass.Iterator
         If .HasLoaded = False Then _
            Call Skynet.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
          
        Call LoadIterator(myclass.Iterator, .Type, .tree, .Root, .Stack.Reverse)
    End With
End Sub

'IPrototype
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericTree: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GenericTree: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.Size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property



