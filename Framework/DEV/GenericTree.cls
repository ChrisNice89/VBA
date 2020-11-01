VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericTree
'*
'* Purpose: A GenericTree is Red-Black tree based navigable map implementation. It provides an efficient means of
'*          storing key-value pairs in sorted order (according to the natural ordering of its keys) without affecting performance
'*          as elements are inserted and deleted.
'*          GenericTree class implements IGenericDictionary interface similar to GenericMap class.
'*          The main difference between them is that GenericMap is an unordered collection while GenericTree is sorted
'*          in the ascending order of its keys. Duplicate Keys are not allowed. A key cannot be null, but a value can be,
'*          if the value type is a reference type.
'*          Changing the sort values of existing items is not supported and may lead to unexpected behavior.
'*          However, if the specified key already exists in the GenericTree, setting the Item[] property overwrites the old value.
'*          In contrast, the Add method throws an exception if an element with the specified key already exists.
'*
'* Notes:   A Red Black Tree guarantee that basic operations like search, get, put and remove take logarithmic time O(log n)
'*
'*
'*======================================================================================================================

'@Folder("<T>Enumerable")

Option Explicit
'@PredeclaredId

Implements IGeneric
Implements IGenericReadOnlyList
Implements IGenericDictionary
Implements IGenericCollection
Implements IGenericIterator

#If Win64 Then
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

Private Const BASE_INDEX As Long = 1

Private Type InternalStack
    Size As Long
    Items() As GenericTreeNode
    Reverse As Boolean
End Type

Private Type Iterator
    Type As DictionaryData
    Current As IGeneric
    Root As GenericTreeNode
    Stack As InternalStack
    HasLoaded As Boolean
End Type

Private Type Member
    Root As GenericTreeNode
    Iterator As Iterator
    FindCounter As Long
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Count() As Long: Count = SizeOf(myclass.Root): End Property

Public Property Set Item(ByVal Key As IGeneric, ByVal Value As IGeneric)

    Dim Node As GenericTreeNode: Set Node = FindNode(Key)
    
    If (Node Is Nothing) Then
         Call Me.Add(Key, Value, True)
    Else
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

Public Function Build() As GenericTree
    
    If (Me Is GenericTree) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption
    
    Set Build = New GenericTree
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
End Sub

Public Function GetMax() As GenericPair
    Set GetMax = FindMaxOf(myclass.Root)
End Function

Public Function GetMin() As GenericPair
    Set GetMin = FindMinOf(myclass.Root)
End Function

Public Function ElementAt(ByVal Index As Long) As GenericPair
    
    If (Index < BASE_INDEX) Or (Index > SizeOf(myclass.Root)) Then _
        Call Skynet.Throw(Me, "Public Function ElementAt").IndexOutofRange("(Index < BASE_INDEX) Or (Index > SizeOf(myclass.Root))")

    Set ElementAt = FindSmallerOrEqual(myclass.Root, Index - BASE_INDEX)
    
End Function

Public Function GetRank(ByVal Key As IGeneric) As Long
    GetRank = RankOf(Key, myclass.Root) + 1
End Function

Public Function GetHigherThan(ByVal Key As IGeneric) As GenericPair
    
    Dim Current As GenericTreeNode: Set Current = myclass.Root
    Dim BestSoFar As GenericTreeNode
    
    Do While (Current Is Nothing) = False
        If Key.CompareTo(Current.Key) = IsLower Then
            Set BestSoFar = Current
            Set Current = Current.Left
        Else
            Set Current = Current.Right
        End If
    Loop
    
    Set GetHigherThan = BestSoFar
    
End Function

Public Function GetLowerThan(ByVal Key As IGeneric) As GenericPair
    
    Dim Current As GenericTreeNode: Set Current = myclass.Root
    Dim BestSoFar As GenericTreeNode
    
    Do While (Current Is Nothing) = False
        If Key.CompareTo(Current.Key) = IsGreater Then
            Set BestSoFar = Current
            Set Current = Current.Right
        Else
            Set Current = Current.Left
        End If
    Loop
    
    Set GetLowerThan = BestSoFar
    
End Function

Private Function RankOf(ByVal Key As IGeneric, ByVal Node As GenericTreeNode) As Long
    
    Dim Order As Long
    
    If Node Is Nothing Then _
        Exit Function
    
    Select Case Key.CompareTo(Node.Key)
        Case Is < 0
            RankOf = RankOf(Key, Node.Left)
        Case Is > 0
            RankOf = 1 + SizeOf(Node.Left) + RankOf(Key, Node.Right)
        Case Else
            RankOf = SizeOf(Node.Left)
    End Select

End Function

Private Function FindSmallerOrEqual(ByVal Node As GenericTreeNode, ByVal Rank As Long) As GenericTreeNode
    
    Dim leftSize As Long
    
    If (Node Is Nothing) Then _
        Set FindSmallerOrEqual = Nothing: Exit Function

    If Node.Left Is Nothing = False Then _
        leftSize = Node.Left.Size
    
    Select Case True
        Case leftSize > Rank
            Set FindSmallerOrEqual = FindSmallerOrEqual(Node.Left, Rank)
        Case leftSize < Rank
            Set FindSmallerOrEqual = FindSmallerOrEqual(Node.Right, Rank - leftSize - 1)
        Case Else
            Set FindSmallerOrEqual = Node
    End Select
    
End Function

Private Function FindMaxOf(ByVal Node As GenericTreeNode) As GenericTreeNode
    
    If Not (Node Is Nothing) Then
        Do While (Node.Right Is Nothing) = False
            Set Node = Node.Right
        Loop
    End If
    
    Set FindMaxOf = Node

End Function

Private Function FindMinOf(ByVal Node As GenericTreeNode) As GenericTreeNode
    
    If Not (Node Is Nothing) Then
        Do While (Node.Left Is Nothing) = False
            Set Node = Node.Left
        Loop
    End If
    
    Set FindMinOf = Node
    
End Function

Private Function InOrderNext(ByVal Node As GenericTreeNode) As GenericTreeNode
    
    Dim P As GenericTreeNode
    
    If Node Is Nothing Then _
        Exit Function
        
    If Node.Right Is Nothing = False Then
        Set P = FindMinOf(Node.Right)
    Else
        Set P = Node.Parent
        Do While (P Is Nothing) = False
            If (P.Right Is Node) = False Then _
                Exit Do
            
            Set Node = P
            Set P = P.Parent
        Loop
    End If
    
    Set InOrderNext = P
        
End Function

Private Function InOrderPrevious(ByVal Node As GenericTreeNode) As GenericTreeNode

    Dim P As GenericTreeNode
    
    If Node Is Nothing Then _
        Exit Function
        
    If (Node.Left Is Nothing) = False Then
        P = FindMaxOf(Node.Left)
    Else
        Set P = Node.Parent
        Do While (P Is Nothing) = False
            If (P.Left Is Node) = False Then _
                Exit Do
            
            Set Node = P
            Set P = P.Parent
        Loop
    End If
    
    Set InOrderPrevious = P

End Function

Public Sub AddAll(ByVal Map As IGenericDictionary)
    
    Dim P As GenericPair

    With Map.Iterator(t:=Pairs_)
        Do While .HasNext(P): Call Me.Add(P.Key, P.Value): Loop
    End With

End Sub

Public Function GetValues(Optional ByVal Reverse As Boolean) As GenericArray
   
    ReDim Values(1 To Me.Count) As IGeneric
    Dim Index As Long
    
    Dim Stack As InternalStack
    Dim Current As GenericTreeNode

    Call LoadStack(Stack, myclass.Root, Reverse)

    Do While InOrderTreeWalk(Current, Stack)
        Set Values(InterlockedIncrement(Index)) = Current.Value
    Loop
    
    Set GetValues = GenericArray.BuildFrom(Values)
    
End Function

Public Function GetKeys(Optional ByVal Reverse As Boolean) As GenericArray
   
    ReDim Keys(1 To Me.Count) As IGeneric
    Dim Index As Long

    Dim Stack As InternalStack
    Dim Current As GenericTreeNode

    Call LoadStack(Stack, myclass.Root, Reverse)

    Do While InOrderTreeWalk(Current, Stack)
        Set Keys(InterlockedIncrement(Index)) = Current.Key
    Loop
    
    Set GetKeys = GenericArray.BuildFrom(Keys)
    
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

    Call LoadIterator(c.Iterator, t, myclass.Root, Reverse)
    Set Iterator = Skynet.CreateInstance(New GenericTree, VarPtr(c), LenB(c))
    
End Function

Public Sub Clear()
    With myclass
        Set .Root = Nothing
    End With
End Sub

Public Function Remove(ByVal Key As IGeneric) As Boolean


End Function

Public Function ContainsKey(ByVal Key As IGeneric) As Boolean: ContainsKey = Not FindNode(Key) Is Nothing: End Function
Public Function ContainsValue(ByVal Value As IGeneric) As Boolean
    
    Dim Stack As InternalStack
    Dim Current As GenericTreeNode

    Call LoadStack(Stack, myclass.Root, False)
    
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
    Dim Order As Long
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function FindNode").KeyIsNothing
    
    Do While (Current Is Nothing) = False
        Order = Key.CompareTo(Current.Key)
    
        If (Order = 0) Then _
            Set FindNode = Current: Exit Function
        
        If Order < 0 Then
            Set Current = Current.Left
        Else
            Set Current = Current.Right
        End If
        
    Loop
    
End Function

Private Function MaxHeightOf(ByVal Root As GenericTreeNode) As Long
    MaxHeightOf = 2 * Logn(SizeOf(Root) + 1, Base:=2)
End Function
Private Function Logn(ByVal Value, Optional ByVal Base As Byte = 2) As Long: Logn = Log(Value) / Log(Base): End Function

Private Function IndexOfValue(ByVal Item As IGeneric) As Long
    
    Dim i As Long
    Dim Stack As InternalStack
    Dim Node As GenericTreeNode: Set Node = myclass.Root
    
    Call LoadStack(Stack, myclass.Root, False)
    
    If Item Is Nothing = False Then
        Do While InOrderTreeWalk(Node, Stack)
            i = i + 1
            If Item.Equals(Node.Value) Then _
                IndexOfValue = i: Exit Function
        Loop
    Else
        Do While InOrderTreeWalk(Node, Stack)
             i = i + 1
            If Node.Value Is Nothing Then _
                IndexOfValue = i: Exit Function
        Loop
    End If
    
    IndexOfValue = -1

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

Public Function Add( _
                        ByVal Key As IGeneric, _
                        ByVal Value As IGeneric, _
                        Optional ByVal Overwrite As Boolean) As Boolean
    
    Dim Node As GenericTreeNode
    Dim Current As GenericTreeNode
    Dim Parent  As GenericTreeNode
    Dim GrandParent As GenericTreeNode
    Dim GreatGrandParent As GenericTreeNode
    Dim Order As Long, i As Long
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Add").KeyIsNothing
    
    If (myclass.Root Is Nothing) Then _
        Set myclass.Root = GenericTreeNode.Build(Key:=Key, Value:=Value, Parent:=Nothing, IsRed:=False): _
        Add = True: _
        Exit Function

    Set Current = myclass.Root
    Do
        Order = Key.CompareTo(Current.Key)
       
        If (Order = 0) Then
            myclass.Root.IsRed = False
            Add = False
            Do While Parent Is Nothing = False
                Parent.Size = Parent.Size - 1
                Set Parent = Parent.Parent
            Loop
            
            If (Overwrite) Then
                Set Current.Value = Value
            Else
                Call Skynet.Throw(Me, "Public Function Add").DuplicateNotAllowed(Key)
            End If
            Exit Function
        End If
        
        Current.Size = Current.Size + 1
        
        If Split4Node(Current) Then
            If (Parent Is Nothing) = False Then
                If Parent.IsRed Then
                    Call EnsureBalance(Current, Parent, GrandParent, GreatGrandParent)
                End If
             End If
        End If
        
        Set GreatGrandParent = GrandParent
        Set GrandParent = Parent
        Set Parent = Current
        
        If Order < 0 Then
            Set Current = Current.Left
        Else
            Set Current = Current.Right
        End If

    Loop While (Current Is Nothing) = False
    
'    If (Parent Is Nothing) Then _
'        Call Skynet.Throw(Me, "Private Sub InsertionBalance").NullPointer("Parent cannot be null here!")

    Set Node = GenericTreeNode.Build(Key:=Key, Value:=Value, Parent:=Parent, IsRed:=True)
    
    If (Order < 0) Then
        Set Parent.Left = Node
    Else
        Set Parent.Right = Node
    End If
    
    If (Parent.IsRed) Then _
        Call EnsureBalance(Node, Parent, GrandParent, GreatGrandParent)
    
    myclass.Root.IsRed = False
    Add = True
     
End Function

Private Sub EnsureBalance( _
                                ByVal Current As GenericTreeNode, _
                                ByRef Parent As GenericTreeNode, _
                                ByVal GrandParent As GenericTreeNode, _
                                ByVal GreatGrandParent As GenericTreeNode)

'    If (GrandParent Is Nothing) Then _
'        Call Skynet.Throw(Me, "Private Sub InsertionBalance").NullPointer("Grand parent cannot be null here!")

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

Private Function SizeOf(ByVal Node As GenericTreeNode) As Long:
    If Node Is Nothing = False Then _
        SizeOf = Node.Size
End Function

Private Sub UpdateSizeOf(ByVal Node As GenericTreeNode)

    Dim i As Long, j As Long
    
    With Node
        If .Left Is Nothing = False Then _
            i = .Left.Size
        
        If .Right Is Nothing = False Then _
            j = .Right.Size
        
        .Size = 1 + i + j
    End With

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

    Dim y  As GenericTreeNode

    With x
        Set y = .Right
        Set .Right = y.Left: Set y.Left = x
        Set y.Parent = .Parent: Set .Parent = y
        Set y.Left.Parent = x
        y.Size = .Size: Call UpdateSizeOf(x)
    End With
    
    Set RotateLeft = y
    
End Function

Private Function RotateRight(ByVal x As GenericTreeNode) As GenericTreeNode

    Dim y  As GenericTreeNode
    
    With x
        Set y = .Left
        Set .Left = y.Right: Set y.Right = x
        Set y.Parent = .Parent: Set .Parent = y
        Set y.Right.Parent = x
        y.Size = .Size: Call UpdateSizeOf(x)
    End With
    
    Set RotateRight = y
    
End Function

Private Function RotateLeftRight(ByVal Node As GenericTreeNode) As GenericTreeNode

'    Dim Child  As GenericTreeNode: Set Child = Node.Left
'    Dim GrandChild  As GenericTreeNode: Set GrandChild = Child.Right
'
'    With GrandChild
'        Set Node.Left = .Right: Set .Right = Node
'        Set Child.Right = .Left: Set .Left = Child
'    End With
'
'    Set RotateLeftRight = GrandChild
    
'    With Child
'        GrandChild.Size = .Size
'        .Size = SizeOf(.Left) + SizeOf(.Right) + 1
'    End With
    
    Set Node.Left = RotateLeft(Node.Left)
    Set RotateLeftRight = RotateRight(Node)

End Function

Private Function RotateRightLeft(ByVal Node As GenericTreeNode) As GenericTreeNode

'    Dim Child  As GenericTreeNode: Set Child = Node.Right
'    Dim GrandChild  As GenericTreeNode: Set GrandChild = Child.Left
'
'    With GrandChild
'        Set Node.Right = .Left: Set .Left = Node
'        Set Child.Left = .Right: Set .Right = Child
'    End With
'
'    Set RotateRightLeft = GrandChild

    Set Node.Right = RotateRight(Node.Right)
    Set RotateRightLeft = RotateLeft(Node)
'
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

Private Sub LoadStack(ByRef Stack As InternalStack, ByVal Node As GenericTreeNode, ByVal Reverse As Boolean)
    
    With Stack
        .Reverse = Reverse
        .Size = 0
        ReDim .Items(1 To MaxHeightOf(Node))
        
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

Private Sub LoadIterator(ByRef e As Iterator, ByVal t As DictionaryData, ByVal Root As GenericTreeNode, ByVal Reverse As Boolean)
    
    If Root Is Nothing Then _
        Call Skynet.Throw(Me, "Private Sub LoadIterator").NullPointer("Root Is Nothing")
    
    With e
        .Type = t
        Set .Current = Nothing
        Set .Root = Root
        Call LoadStack(.Stack, .Root, Reverse)
        .HasLoaded = True
    End With
    
End Sub

Private Function CopyOf(ByVal t As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long, ByVal Caller As String) As GenericArray
    
    If (GArray Is Nothing) Then _
        Call Skynet.Throw(Me, Caller).NullPointer("GArray Is Nothing")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call Skynet.Throw(Me, Caller).IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < SizeOf(myclass.Root)) Then _
        Call Skynet.Throw(Me, Caller).InvalidInput("(GArray.Length - Index < myClass.Count)")
    
    Dim Stack As InternalStack
    Dim Node As GenericTreeNode
    Dim Current As IGeneric
    Call LoadStack(Stack, myclass.Root, False)

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
                Set GArray(ArrayIndex) = Node 'GenericPair
                ArrayIndex = ArrayIndex + 1
            Loop
    End Select

    Set CopyOf = GArray
    
End Function

'IGeneric
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GenericTree: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = Virtual.Pointer: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = Virtual.Size: End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericTree):  End Property
Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member
    
    With c
        Set .Root = myclass.Root
    End With
    
    Set IGeneric_Clone = Skynet.CreateInstance(New GenericTree, VarPtr(c), LenB(c))
    
End Function

Private Sub IGeneric_Dispose()
    
    If myclass.Root Is Nothing = False Then
        With myclass.Root
            If .Left Is Nothing = False Then _
                Call Skynet.Dispose(.Left)
        
            If .Right Is Nothing = False Then _
                Call Skynet.Dispose(.Right)
        
            Set .Parent = Nothing
            Set .Left = Nothing
            Set .Right = Nothing
        End With
    End If
    
    myclass = EmptyStructure
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    Call Skynet.Throw(Me, "Private Function IGeneric_Equals").NotImplemented
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    Call Skynet.Throw(Me, "Private Function IGeneric_Equals").NotImplemented
End Function

Private Function IGeneric_ToString() As String
    With myclass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Elements: " & Me.Count
    End With
End Function

Private Sub Class_Terminate()
    If Not myclass.Root Is Nothing Then _
        Call Skynet.Dispose(myclass.Root)
    
    myclass = EmptyStructure
    
End Sub

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_BaseIndex() As Long: IGenericReadOnlyList_BaseIndex = BASE_INDEX: End Property
Private Property Get IGenericReadOnlyList_Count() As Long: IGenericReadOnlyList_Count = Me.Count: End Property
Private Property Get IGenericReadOnlyList_ElementAt(ByVal Index As Long) As IGeneric: Set IGenericReadOnlyList_ElementAt = Me.ElementAt(Index): End Property

Private Sub IGenericReadOnlyList_Clear(): Call Me.Clear: End Sub
Private Sub IGenericReadOnlyList_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(t:=Values_, GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="IGenericReadOnlyList_CopyTo"):  End Sub
Private Function IGenericReadOnlyList_IndexOf(ByVal Item As IGeneric) As Long: IGenericReadOnlyList_IndexOf = IndexOfValue(Item): End Function
Private Function IGenericReadOnlyList_Iterator() As IGenericIterator: Set IGenericReadOnlyList_Iterator = Me.Iterator(t:=Values_): End Function

'IGenericDictionary
Private Property Set IGenericDictionary_Item(ByVal Key As IGeneric, ByVal Value As IGeneric): Set Me.Item(Key) = Value: End Property
Private Property Get IGenericDictionary_Item(ByVal Key As IGeneric) As IGeneric: Set IGenericDictionary_Item = Me.Item(Key): End Property
Private Property Get IGenericDictionary_Count() As Long: IGenericDictionary_Count = Me.Count: End Property
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
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = False: End Property
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.ContainsKey(Key:=Item): End Function
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = CopyOf(t:=Pairs_, GArray:=GenericArray.Build(Me.Count), ArrayIndex:=GenericArray.LowerBound, Caller:="Private Function IGenericCollection_ToArray() As GenericArray"): End Function
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
                Set .Current = Node 'GenericPair
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
          
        Call LoadIterator(myclass.Iterator, .Type, .Root, .Stack.Reverse)
    End With
End Sub
