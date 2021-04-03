VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSortedSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericSortedSet
'*
'* Purpose: A GenericSortedSet is Red-Black tree based implementation. It provides an efficient means of
'*          storing elements  in sorted order (according to the natural ordering of the elements) without affecting performance
'*          as elements are inserted and deleted.
'*          Duplicate elements are not allowed. A element can be null
'*          Changing the sort values of existing Elements is not supported and may lead to unexpected behavior.
'*          If the specified element already exists in the GenericSortedSet,
'*          the Add method will not throw an exception if an element with the specified value already exists.
'*          We can quickly find the ith smallest number in a set or the rank of a given element in the total ordering of the set.
'*
'*          RB-trees are binary trees satisfying some additional constraints:
'*              1. Every node is colored either red or black. (Equivalently, the edge from the node to its parent is colored red or black.)
'*              2. Leaf nodes (null pointers) are black.
'*              3. No consecutive reds - if a node is red, its parent is black.
'*              4. Every leaf (null) has the same "black depth" - the number of black nodes on the path from the leaf to the root.
'*
'* Notes:   A Red Black Tree guarantee that basic operations like search, get, put and remove take logarithmic time O(log n)
'*          It is generally used when you have to store unique elements and maintain an (ascending) order.
'*          It maintains (ascending) order (according to the associated Comparer implementation) and does not store duplicate elements.
'*          Add method returned a boolean that indicates if it was successfully added.
'*          This tree implementation supports general order-statistic operations.
'*          Every node knows the number of items in its (left) subtree.  This permits O(log n) implementations of search-by-index,
'*          IndexOf(x), and other index-based operations.
'*
'*          A left-leaning tree has one additional constraint:
'*              5. A mixed-color family leans left.   If a node has one red child and one black child, the red child is the left child.
'*
'*          This constraint makes the rebalancing algorithms simpler, as there are fewer cases to worry about.
'*
'*
'*======================================================================================================================

Implements IGeneric
Implements IGenericReadOnlyList
Implements IGenericList
Implements IGenericCollection
Implements IGenericIterator

#If Win64 Then
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

Private Const BASE_INDEX As Long = 0
Private Const SORT_ORDER As Long = 1 'Ascending
Private Const IS_RED As Byte = 1
Private Const IS_BLACK As Byte = 0

Private Type InternalStack
    Size As Long
    Elements() As GenericNode
    Reverse As Boolean
End Type

Private Type Iterator
    Current As IGeneric
    Root As GenericNode
    Stack As InternalStack
    HasLoaded As Boolean
End Type

Private Type Member
    Root As GenericNode
    Comparer As IGenericComparer
    Iterator As Iterator
End Type
Private myclass As Member

Public Property Get Elements() As IGenericCollection: Set Elements = Me: End Property
Public Property Get First() As Long: First = BASE_INDEX: End Property
Public Property Get Last() As Long
    
    Dim c As Long
    Last = BASE_INDEX
    c = SizeOf(myclass.Root)
    
    If c > 0 Then _
        Last = c - Last - 1
   
End Property
Public Property Get SortOrder() As SortOrder: SortOrder = SORT_ORDER: End Property
Private Property Get IsHeightBalanced(): IsHeightBalanced = CheckHeightBalance(myclass.Root): End Property
Private Property Get IsBalanced(): IsBalanced = CheckRedBlackBalance(myclass.Root): End Property
Public Property Get Comparer() As IGenericComparer: Set Comparer = myclass.Comparer: End Property

Public Function Build(Optional ByVal Comparer As IGenericComparer) As GenericSortedSet
        
    Dim newSet As Member
    
    If Comparer Is Nothing Then _
        Set Comparer = IGenericComparer
    
    Set newSet.Comparer = Comparer
    Set Build = System.CreateInstance(New GenericSortedSet, VarPtr(newSet), LenB(newSet))
    
End Function

Public Function Of(ByVal Comparer As IGenericComparer, ParamArray GenericValues() As Variant) As IGenericReadOnlyList
        
    Dim SortedSet As GenericSortedSet
    Dim i As Long
    
    If IsMissing(GenericValues) Then _
        Call System.Throw(Me, "Public Function Of").NoSuchElement("ParamArray GenericValues are missing")
        
    If Comparer Is Nothing Then _
        Set Comparer = IGenericComparer
    
    Set SortedSet = GenericSortedSet.Build(Comparer)
    
    With SortedSet
        For i = LBound(GenericValues) To UBound(GenericValues)
            
            If Not VBA.IsObject(GenericValues(i)) Then _
                Call System.Throw(Me, "Public Function Of").InvalidInput("Datatype {" & TypeName$(GenericValues(i)) & "} not allowed")
            
            If (GenericValues(i) Is Nothing) Then _
                Call System.Throw(Me, "Public Function Of").InvalidInput("Element is nothing")
            
            If Not TypeOf GenericValues(i) Is IGeneric Then _
                Call System.Throw(Me, "Public Function Of").InvalidInput("Not TypeOf GenericValues(i) Is IGeneric")
                
            Call .Add(Value:=GenericValues(i))
            
        Next
    End With
    
    Set Of = SortedSet
    
End Function

Public Function GetMax() As IGeneric: Set GetMax = FindMaxOf(myclass.Root): End Function
Public Function GetMin() As IGeneric: Set GetMin = FindMinOf(myclass.Root): End Function

Public Function RankOf(ByVal Value As IGeneric) As Long
    
    Dim Node As GenericNode
    Dim Order As CompareResult
    
    Set Node = myclass.Root
    
    Do While Node Is Nothing = False
        Order = myclass.Comparer.Compare(Value, Node.Value) '* SORT_ORDER
        
        If (Order = IsEqual) Then _
            RankOf = RankOf + SizeOf(Node.Left): Exit Function
            
        If Order = islower Then
            Set Node = Node.Left
        Else
            RankOf = RankOf + 1 + SizeOf(Node.Left): Set Node = Node.Right
        End If
    Loop

End Function

Public Function GetAt(ByVal Index As Long) As IGeneric
Attribute GetAt.VB_UserMemId = 0

    If (Index >= SizeOf(myclass.Root)) Then _
        Call System.Throw(Me, "Public Function GetAt").IndexOutofRange("(Index >= SizeOf(myclass.Root)")
    
    Set GetAt = FindNodeAt(Index).Value
    
End Function

Public Function IndexOf(ByVal Value As IGeneric) As Long
    
    Dim Current As GenericNode: Set Current = myclass.Root
    Dim Order As CompareResult
    
    Do While Current Is Nothing = False
        Order = myclass.Comparer.Compare(Value, Current.Value) '* SORT_ORDER
        
        If (Order = IsEqual) Then _
            IndexOf = IndexOf + SizeOf(Current.Left): Exit Function
            
        If Order = islower Then
            Set Current = Current.Left
        Else
            IndexOf = IndexOf + 1 + SizeOf(Current.Left): Set Current = Current.Right
        End If
    Loop
   
    IndexOf = -1
        
End Function

Public Function GetHigherThan(ByVal Value As IGeneric) As IGeneric
    
    Dim Current As GenericNode: Set Current = myclass.Root
    Dim BestSoFar As GenericNode
    
    With myclass.Comparer
        Do While (Current Is Nothing) = False
            If .Compare(Value, Current.Value) = islower Then '* SORT_ORDER
                Set BestSoFar = Current
                Set Current = Current.Left
            Else
                Set Current = Current.Right
            End If
        Loop
    End With
    
    Set GetHigherThan = BestSoFar.Value
    
End Function

Public Function GetLowerThan(ByVal Value As IGeneric) As IGeneric
    
    Dim Current As GenericNode: Set Current = myclass.Root
    Dim BestSoFar As GenericNode
    
    With myclass.Comparer
        Do While (Current Is Nothing) = False
            If .Compare(Value, Current.Value) = isgreater Then
                Set BestSoFar = Current
                Set Current = Current.Right
            Else
                Set Current = Current.Left
            End If
        Loop
    End With
    
    Set GetLowerThan = BestSoFar.Value
    
End Function

Public Function Reverse() As IGenericIterator

    Dim c As Member

    Call LoadIterator(c.Iterator, myclass.Root, Reverse:=True)
    Set Reverse = System.CreateInstance(New GenericSortedSet, VarPtr(c), LenB(c))

End Function

Public Sub CopyTo( _
                    ByVal GArray As GenericArray, _
                    ByVal ArrayIndex As Long, _
                    Optional ByVal Index As Long = BASE_INDEX, _
                    Optional ByVal Count As Long = 0)
    
    If (GArray Is Nothing) Then _
        Call System.Throw(Me, "Public Function CopyTo").IsNothing("(GArray Is Nothing)")
    
    If (GArray.IsEmpty) Then _
        Call System.Throw(Me, "Public Function CopyTo").IsNothing("(GArray.IsEmpty)")
        
    If (myclass.Root Is Nothing) Then _
        Call System.Throw(Me, "Public Function CopyTo").IsNothing("(myclass.Root Is Nothing)")
        
    If (Count = 0) Then _
        Count = myclass.Root.Count - Index
        
    If (myclass.Root.Count - Index < Count) Then _
        Call System.Throw(Me, "Public Sub CopyTo").ArgOutOfRange("(Me.Count - Index < Count)")
    
    Dim Stack As InternalStack
    Dim Current As GenericNode
    
    If (Index = BASE_INDEX) Then
        Set Current = myclass.Root
    Else
        Set Current = FindNodeAt(Index)
    End If
    
    ArrayIndex = ArrayIndex - 1
    
    Do While InterlockedIncrement(Index) <= Count
         With GArray
            Set .ElementAt(InterlockedIncrement(ArrayIndex)) = Current.Value
            Set Current = InOrderNext(Current)
        End With
    Loop
    
End Sub

Public Function DoUnionWith(ByVal Collection As IGenericCollection) As GenericSortedSet

    Dim Enumerable As IGenericIterator
    Dim Element As IGeneric
    
    If (Collection Is Nothing) Then _
        Call System.Throw(Me, "Public Sub AddAll").IsNothing("(Collection Is Nothing)")
    
    If (Collection Is Me) Then _
        Exit Function
    
    If Collection.Count > 0 Then
        Set Enumerable = Collection.Iterator
    Else
        Set Enumerable = Collection
    End If
    
    With Enumerable
        Do While .HasNext(Element)
'            If FindNode(Element) Is Nothing Then
                Call Me.Add(Element)
        Loop
    End With
    
    Set DoUnionWith = Me
    
End Function

Public Function DoExeceptWith(ByVal Collection As IGenericCollection) As Long
    
    Call System.Throw(Me, "Public Function ExeceptWith").NotImplemented("Not implemented yet")
        
    Dim Candidate As IGeneric
    Dim Result As Long
    Dim Min As GenericNode
    Dim Max As GenericNode
    Dim Node As GenericNode

    If (Collection Is Nothing) Then _
        Call System.Throw(Me, "Public Function DoExeceptWith").IsNothing("(Collection Is Nothing)")
    
    If Me.Elements.Count = 0 Then _
        Exit Function
        
    If (Collection Is Me) Then _
        Call Me.Elements.Clear: Exit Function
    
    If (Collection.Count > 0) Then
    
        Set Min = Me.GetMin
        Set Max = Me.GetMax
    
        With Collection.Iterator
            Do While .HasNext(Candidate)
                If Not (myclass.Comparer.Compare(Candidate, Min) = islower) Then
                    If Not (myclass.Comparer.Compare(Candidate, Max) = isgreater) Then
                        If RemoveNode(FindNode(Candidate)) Then _
                            Result = Result + 1
                    End If
                End If
            Loop
        End With
    End If
    
    DoExeceptWith = Result
    
End Function

Public Function DoIntersectWith(ByVal Collection As IGenericCollection) As Long

    Dim Value As IGeneric
    Dim NodesToSave As GenericOrderedList
    Dim Node As GenericNode
    Dim Current As GenericNode
    Dim Parent  As GenericNode
    Dim GrandParent As GenericNode
    Dim GreatGrandParent As GenericNode
    Dim Order As CompareResult
    
    If Collection Is Nothing Then
    'Error
    End If
    
    If Collection.Count = 0 Then
        Exit Function
    End If
    
    If IGeneric_IsRelatedTo(Collection) Then
        'Special handling:exit sub
    End If
    
    Set NodesToSave = GenericOrderedList.Build(Me.Elements.Count)
    
    With Collection.Iterator
    
        Do While .HasNext(Value)
            Set Node = FindNode(Value)
            'If RemoveNode(Node) Then
            If Not Node Is Nothing Then
                With Node 'reset
                    '  Set .Value = Value 'not necessary -if not reference equal?!
                    .Count = 1
                    .State = IS_BLACK
                End With
                Call NodesToSave.Add(Node)
            End If
        Loop
        
    End With
    
    Call Me.Elements.Clear
    
    With NodesToSave.Elements.Iterator
    
        If .HasNext(Node) Then _
            Set myclass.Root = Node
        
        Do While .HasNext(Node)
            Set Current = myclass.Root
            Set Value = Node.Value
            
            With myclass.Comparer
            
                Do
                    Order = .Compare(Value, Current.Value) '*SORT_ORDER
        
                    Current.Count = Current.Count + 1
                    
                    If Split4Node(Current) Then _
                        If (Parent Is Nothing) = False Then _
                            If Parent.State = IS_RED Then _
                                Call EnsureBalance(Current, Parent, GrandParent, GreatGrandParent)
                    
                    Set GreatGrandParent = GrandParent
                    Set GrandParent = Parent
                    Set Parent = Current
                    
                    If Order = islower Then
                        Set Current = Current.Left
                    Else
                        Set Current = Current.Right
                    End If
            
                Loop While (Current Is Nothing) = False
                
            End With
        '    If (Parent Is Nothing) Then _
        '        Call System.Throw(Me, "Private Sub InsertionBalance").NullPointer("Parent cannot be null here!")
            
            If (Order = islower) Then
                Set Parent.Left = Node
            Else
                Set Parent.Right = Node
            End If
            
            Set Node.Parent = Parent
            Node.State = IS_RED
            
            If (Parent.State = IS_RED) Then _
                Call EnsureBalance(Node, Parent, GrandParent, GreatGrandParent)
            
            myclass.Root.State = IS_BLACK
        Loop
        
    End With
    
End Function

Public Function Remove(ByVal Value As IGeneric) As Boolean
    Call System.Throw(Me, "Public Function Remove").NotImplemented("Not implemented yet")
    Remove = RemoveNode(FindNode(Value))
End Function

Public Function Contains(ByVal Value As IGeneric) As Boolean: Contains = Not FindNode(Value:=Value) Is Nothing: End Function

Public Function Add(ByVal Value As IGeneric) As Boolean
    
    Dim Node As GenericNode
    Dim Current As GenericNode
    Dim Parent  As GenericNode
    Dim GrandParent As GenericNode
    Dim GreatGrandParent As GenericNode
    Dim Order As CompareResult

    If (myclass.Root Is Nothing) Then
        Set myclass.Root = New GenericNode
        With myclass.Root
            Set .Value = Value
            .State = IS_BLACK
            .Count = 1
        End With
        Add = True
        Exit Function
    End If
    
    Set Current = myclass.Root
    With myclass.Comparer
        Do
            Order = .Compare(Value, Current.Value) '*SORT_ORDER
           
            If (Order = IsEqual) Then
                myclass.Root.State = IS_BLACK
                Do While Parent Is Nothing = False
                    Call Parent.AddCount(-1)
                    Set Parent = Parent.Parent
                Loop
                Exit Function
            End If
            
            Call Current.AddCount(1)
            
            If Split4Node(Current) Then _
                If (Parent Is Nothing) = False Then _
                    If Parent.State = IS_RED Then _
                        Call EnsureBalance(Current, Parent, GrandParent, GreatGrandParent)
            
            Set GreatGrandParent = GrandParent
            Set GrandParent = Parent
            Set Parent = Current
            
            If Order = islower Then
                Set Current = Current.Left
            Else
                Set Current = Current.Right
            End If
    
        Loop While (Current Is Nothing) = False
    End With
'    If (Parent Is Nothing) Then _
'        Call System.Throw(Me, "Private Sub InsertionBalance").NullPointer("Parent cannot be null here!")

    Set Node = New GenericNode
    
    With Node
        Set .Value = Value
        Set .Parent = Parent
        .State = IS_RED
        .Count = 1
    End With
    
    If (Order = islower) Then
        Set Parent.Left = Node
    Else
        Set Parent.Right = Node
    End If
    
    If (Parent.State = IS_RED) Then _
        Call EnsureBalance(Node, Parent, GrandParent, GreatGrandParent)
    
    myclass.Root.State = IS_BLACK
    Add = True
     
End Function

'Public Sub AddLeftLeaning(ByVal Value As IGeneric)
'    Set myclass.Root = InsertLeftLeaning(myclass.Root, Value)
'    myclass.Root.IsRed = False
'End Sub

Private Function FindNodeAt(ByVal Rank As Long) As GenericNode
    
    Dim leftSize As Long
    
    Dim Node As GenericNode
    Set Node = myclass.Root
    
    Do While Node Is Nothing = False
        If Node.Left Is Nothing = False Then
            leftSize = Node.Left.Count
        Else
            leftSize = 0
        End If
        
        If Rank = leftSize Then _
            Exit Do
      
        If (leftSize > Rank) Then
            Set Node = Node.Left
        Else
            Set Node = Node.Right
            Rank = Rank - leftSize - 1
        End If
    Loop
    
    Set FindNodeAt = Node
    
'    Set Result = Node
'
'    i = Node.Count
'
'    Do While (i - 1 > Rank)
'        i = ((i - 1) / 2)
'        If (Rank < i) Then
'            Set Result = Result.Left
'        Else
'            Set Result = Result.Right
'            Rank = Rank - i
'        End If
'    Loop
'
'    Set SelectKth = IIf(i - 1 = Rank, Result, Nothing)


'    Set SelectKth = IIf(Rank = leftSize, Node, Nothing)
    
'    Node kSmallest(Node root, int k) {
'  int i = root.size(); // 2^height - 1, single node is height = 1;
'  Node result = root;
'  while (i - 1 > k) {
'    i = (i-1)/2;  // size of left subtree
'    if (k < i) {
'      result = result.left;
'    } else {
'      result = result.right;
'      k -= i;
'    }
'  }
'  return i-1==k ? result: null;
'}
'    If (Node Is Nothing) Then _
'        Set SelectKth = Nothing: Exit Function
'
'
'    If Node.Left Is Nothing = False Then _
'        leftSize = Node.Left.Count
'
'
'    Select Case True
'        Case leftSize > Rank
'            Set SelectKth = SelectKth(Node.Left, Rank)
'        Case leftSize < Rank
'            Set SelectKth = SelectKth(Node.Right, Rank - leftSize - 1)
'        Case Else
'            Set SelectKth = Node
'    End Select
    
End Function

Private Function FindMaxOf(ByVal Node As GenericNode) As GenericNode
    
    If Not (Node Is Nothing) Then _
        Do While (Node.Right Is Nothing) = False: Set Node = Node.Right: Loop
    
    Set FindMaxOf = Node

End Function

Private Function FindMinOf(ByVal Node As GenericNode) As GenericNode
    
    If Not (Node Is Nothing) Then _
        Do While (Node.Left Is Nothing) = False: Set Node = Node.Left: Loop
    
    Set FindMinOf = Node
    
End Function

Private Function InOrderNext(ByVal Node As GenericNode) As GenericNode
    
    Dim P As GenericNode
    
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

Private Function InOrderPrevious(ByVal Node As GenericNode) As GenericNode

    Dim P As GenericNode
    
    If Node Is Nothing Then _
        Exit Function
        
    If (Node.Left Is Nothing) = False Then
        Set P = FindMaxOf(Node.Left)
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

Private Function FindNode(ByVal Value As IGeneric) As GenericNode
    
    Dim Current As GenericNode: Set Current = myclass.Root
    Dim Order As CompareResult
    
    With myclass.Comparer
        Do While (Current Is Nothing) = False
        
            Order = .Compare(Value, Current.Value) '*SORT_ORDER
           
            If (Order = IsEqual) Then _
                Set FindNode = Current: Exit Function
            
            If Order = islower Then
                Set Current = Current.Left
            Else
                Set Current = Current.Right
            End If
            
        Loop
    End With
    
End Function

'Todo - test performance vs stack ' Node.Count ?!
'Use like a Queue
Private Sub ClearTree(ByRef Node As GenericNode)

    Dim Left As GenericNode
    Dim Right As GenericNode
    
    Do While Node Is Nothing = False
        With Node
            Set .Parent = Nothing
            Set Left = .Left
            
            If Not Left Is Nothing Then
                Set .Left = Left.Right '// Rotate the left child up.
                Set Left.Right = Node
                Set Node = Left
            Else
                Set Right = .Right
                Set .Right = Nothing
                Set Node = Right
            End If
        End With
    Loop
 
End Sub

Private Function RemoveNode(ByVal Node As GenericNode) As Boolean

    Call System.Throw(Me, "Private Sub RemoveNode").NotImplemented("Not implemented yet")
    
    If (Node Is Nothing) Then _
        Exit Function
    
    
    With Node
        Set .Parent = Nothing
        Set .Left = Nothing
        Set .Right = Nothing
    End With
    
    RemoveNode = True
    
End Function

Private Function MaxHeightOf(ByVal Root As GenericNode) As Long: MaxHeightOf = 2 * LOGn(SizeOf(Root) + 1, Base:=2): End Function
Private Function LOGn(ByVal Value, Optional ByVal Base As Byte = 2) As Long: LOGn = Log(Value) / Log(Base): End Function

Private Sub ReplaceNode( _
                            ByVal Match As GenericNode, _
                            ByVal ParentOfMatch As GenericNode, _
                            ByVal Succesor As GenericNode, _
                            ByVal ParentOfSuccesor As GenericNode)

    If (Succesor Is Match) Then
        If Not (Match.Right Is Nothing) Then _
            Call System.Throw(Me, "Private Sub ReplaceNode").InvalidOperation("Right child must be null!")

        Set Succesor = Match.Left
    Else
    
        If Not (Succesor.Right Is Nothing) Then _
            Succesor.Right.State = IS_RED
        
        If Not (ParentOfSuccesor Is Match) Then _
            Set ParentOfSuccesor.Left = Succesor.Right: _
            Set Succesor.Right = Match.Right
            
        Set Succesor.Left = Match.Left

    End If
        
    If Not (Succesor Is Nothing) Then _
        Succesor.State = Match.State
                
    Call ReplaceChildOfNodeOrRoot(ParentOfMatch, Match, Succesor)
                   
End Sub
'leftleaning
Private Function InsertLeftLeaning(ByVal Node As GenericNode, ByVal Value As IGeneric) As GenericNode
    
    If Node Is Nothing Then
        Set InsertLeftLeaning = New GenericNode
        
        With InsertLeftLeaning
            Set .Value = Value
'            Set .Parent = Node.Parent
            .State = IS_RED
            .Count = 1
        End With
        Exit Function
    End If
    
    With Node
        '.Count = .Count + 1
        Select Case myclass.Comparer.Compare(Value, .Value) '*SORT_ORDER
            Case Is < CompareResult.islower
                Set .Left = InsertLeftLeaning(.Left, Value)
            Case Is > CompareResult.isgreater
                Set .Right = InsertLeftLeaning(.Right, Value)
            Case Else
                Set .Value = Value
        End Select
        .Count = .Count + 1
    End With
    
    If (IsRed(Node.Right)) Then _
        Set Node = RotateLeft(Node): Node.State = Node.Left.State: Node.Left.State = IS_RED
    
    If (IsRed(Node.Left)) Then _
        If IsRed(Node.Left.Left) Then _
            Set Node = RotateRight(Node): Node.State = Node.Right.State: Node.Right.State = IS_RED
    
    With Node
        If (IsRed(.Left)) Then _
            If IsRed(.Right) Then _
                Call FlipColor(Node)
    End With
    
    Set InsertLeftLeaning = Node
    
End Function

Private Sub EnsureBalance( _
                                ByVal Current As GenericNode, _
                                ByRef Parent As GenericNode, _
                                ByVal GrandParent As GenericNode, _
                                ByVal GreatGrandParent As GenericNode)

'    If (GrandParent Is Nothing) Then _
'        Call System.Throw(Me, "Private Sub InsertionBalance").NullPointer("Grand parent cannot be null here!")

    Dim CurrentIsOnRight As Boolean: CurrentIsOnRight = Parent.Right Is Current
    Dim NewChildOfGreatGrandParent As GenericNode
    
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
    
    GrandParent.State = IS_RED
    NewChildOfGreatGrandParent.State = IS_BLACK
    
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

'Recursive
Private Sub AddToArray(ByVal Node As GenericNode, ByVal GArray As GenericArray, ByRef Index As Long)

    If Not (Node.Left Is Nothing) Then _
        Call AddToArray(Node.Left, GArray, Index)
    
    Set GArray(Index) = Node.Value: Index = Index + 1
    
    If Not (Node.Right Is Nothing) Then _
        Call AddToArray(Node.Right, GArray, Index)
    
End Sub

Private Function CheckRedBlackBalance(ByVal Node As GenericNode, Optional ByRef maxH As Long, Optional ByRef minH As Long) As Boolean
    
    Dim lmaxH As Long, lminH As Long
    Dim rmaxH As Long, rminH As Long
    
    If Node Is Nothing Then
        maxH = 0
        minH = 0
        CheckRedBlackBalance = True
        Exit Function
    End If

    If Not CheckRedBlackBalance(Node.Left, lmaxH, lminH) Then _
        Exit Function
        
    If Not CheckRedBlackBalance(Node.Right, rmaxH, rminH) Then _
        Exit Function

    maxH = IIf(lmaxH < rmaxH, rmaxH, lmaxH) + 1
    minH = IIf(lminH > rminH, rminH, lminH) + 1
    
    If maxH <= 2 * minH Then _
        CheckRedBlackBalance = True
    
End Function

Private Function CheckHeightBalance(ByVal Node As GenericNode) As Boolean

    Dim lh As Long
    Dim rh As Long
    
    If Node Is Nothing Then _
        CheckHeightBalance = True: Exit Function
        
    lh = HeightOf(Node.Left)
    rh = HeightOf(Node.Right)
    
    'High balanced
    If Abs(lh - rh) <= 1 Then _
        If CheckHeightBalance(Node.Left) And CheckHeightBalance(Node.Left) Then _
            CheckHeightBalance = True
  
End Function

Private Function HeightOf(ByVal Node As GenericNode) As Long
    
    Dim l As Long
    Dim r As Long
    
    If Not (Node Is Nothing) Then
        l = HeightOf(Node.Left)
        r = HeightOf(Node.Right)
        HeightOf = IIf(l > r, l, r) + 1
    Else
        HeightOf = -1
    End If

End Function

Private Sub ReplaceChildOfNodeOrRoot( _
                                        ByVal Parent As GenericNode, _
                                        ByVal Child As GenericNode, _
                                        ByVal NewChild As GenericNode)
    
    If Parent Is Nothing Then _
        Set myclass.Root = NewChild: Exit Sub
    
    If (Parent.Left Is Child) Then
        Set Parent.Left = NewChild
    Else
        Set Parent.Right = NewChild
    End If
        
End Sub

Private Function SizeOf(ByVal Node As GenericNode) As Long:
    If Node Is Nothing = False Then _
        SizeOf = Node.Count
End Function

Private Function Is4Node(ByVal Node As GenericNode) As Boolean
    If (IsRed(Node.Left)) Then _
        Is4Node = IsRed(Node.Right)
End Function

Private Function IsRed(ByVal Node As GenericNode) As Boolean
    If Not (Node Is Nothing) Then _
        IsRed = Node.State = IS_RED
End Function

Private Function IsBlack(ByVal Node As GenericNode) As Boolean
    If Not (Node Is Nothing) Then _
        IsBlack = Node.State = IS_BLACK
End Function

Private Function IsNothingOrBlack(ByVal Node As GenericNode) As Boolean
    If (Node Is Nothing) Then _
        IsNothingOrBlack = True: Exit Function
    
    IsNothingOrBlack = Node.State = IS_BLACK
End Function

Private Function Split4Node(ByVal Node As GenericNode) As Boolean
    With Node
        If (IsRed(.Left)) Then
            If IsRed(.Right) Then
                .State = IS_RED
                .Left.State = IS_BLACK
                .Right.State = IS_BLACK
                Split4Node = True
            End If
        End If
    End With
End Function

Private Sub FlipColor(ByVal Node As GenericNode)
    
    If Not Node Is Nothing Then
        With Node
            .State = IIf(.State = IS_RED, IS_BLACK, IS_RED)
            .Left.State = IIf(.State = IS_RED, IS_BLACK, IS_RED)
            .Right.State = IIf(.State = IS_RED, IS_BLACK, IS_RED)
        End With
    End If

End Sub

Private Function RotateLeft(ByVal X As GenericNode) As GenericNode

    Dim Y  As GenericNode
    Dim i As Long, j As Long
    
    With X
        Set Y = .Right
        Set .Right = Y.Left
        
        If Not .Right Is Nothing Then _
            Set .Right.Parent = X: j = .Right.Count
        
        Set Y.Parent = .Parent
        Set Y.Left = X: Set .Parent = Y
        
        Y.Count = .Count
        
        If Not .Left Is Nothing Then _
            i = .Left.Count

        .Count = 1 + i + j
    End With
    
    Set RotateLeft = Y
    
End Function

Private Function RotateRight(ByVal X As GenericNode) As GenericNode

    Dim Y  As GenericNode
    Dim i As Long, j As Long
    
    With X
        Set Y = .Left
        Set .Left = Y.Right
        
        If Not .Left Is Nothing Then _
            Set .Left.Parent = X: i = .Left.Count
        
        Set Y.Parent = .Parent
        Set Y.Right = X: Set .Parent = Y
        
        Y.Count = .Count
        
        If Not .Right Is Nothing Then _
           j = .Right.Count
            
        .Count = 1 + i + j
    End With

    Set RotateRight = Y
    
End Function

Private Function RotateLeftRight(ByVal Node As GenericNode) As GenericNode
    
    Set Node.Left = RotateLeft(Node.Left)
    Set RotateLeftRight = RotateRight(Node)

End Function

Private Function RotateRightLeft(ByVal Node As GenericNode) As GenericNode

    Set Node.Right = RotateRight(Node.Right)
    Set RotateRightLeft = RotateLeft(Node)

End Function

Private Sub UpdateSizeOf(ByVal Node As GenericNode)

    Dim i As Long, j As Long
    Dim leftNode As GenericNode, rightNode As GenericNode
    
    With Node
        Set leftNode = .Left
        
        If leftNode Is Nothing = False Then _
            i = leftNode.Count
        
        Set rightNode = .Right
        If rightNode Is Nothing = False Then _
            j = rightNode.Count
            
        .Count = 1 + i + j
    End With

End Sub

Private Function InOrderTreeWalk( _
                                    ByRef Current As GenericNode, _
                                    ByRef Stack As InternalStack) As Boolean
    
    Dim Node As GenericNode

    If Stack.Size = 0 Then _
        Set Current = Nothing: Exit Function
    
    With Stack
        Set Current = .Elements(.Size)
        Set .Elements(.Size) = Nothing
        .Size = .Size - 1
        If .Reverse = False Then
            Set Node = Current.Right
    
            Do While Not (Node Is Nothing)
                Set .Elements(InterlockedIncrement(.Size)) = Node
                Set Node = Node.Left
            Loop
        Else
            Set Node = Current.Left
    
            Do While Not (Node Is Nothing)
                Set .Elements(InterlockedIncrement(.Size)) = Node
                Set Node = Node.Right
            Loop
        End If
    End With
    
    InOrderTreeWalk = True

End Function

Private Sub LoadStack(ByRef Stack As InternalStack, ByVal Node As GenericNode, ByVal Reverse As Boolean)
    
    With Stack
        .Reverse = Reverse
        .Size = 0
        ReDim .Elements(1 To MaxHeightOf(Node))
        
        If .Reverse = False Then
            Do While Not (Node Is Nothing)
                Set .Elements(InterlockedIncrement(.Size)) = Node
                Set Node = Node.Left
            Loop
        Else
            Do While Not (Node Is Nothing)
                Set .Elements(InterlockedIncrement(.Size)) = Node
                Set Node = Node.Right
            Loop
        End If
    End With
    
End Sub

Private Sub LoadIterator(ByRef e As Iterator, ByVal Root As GenericNode, ByVal Reverse As Boolean)
    
    If (Root Is Nothing) Then _
        Call System.Throw(Me, "Private Sub LoadIterator").IsNothing("(Root Is Nothing)")
    
    With e
        Set .Root = Root
        Call LoadStack(.Stack, .Root, Reverse)
        .HasLoaded = True
    End With
    
End Sub

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericSortedSet: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GenericSortedSet: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericSortedSet):  End Property
Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = Me.Elements.Copy: End Function

Private Function CloneTree(ByVal Node As GenericNode) As GenericNode
    
    Dim Clone As GenericNode
    
    If Node Is Nothing Then _
        Exit Function
    Set Clone = Node.Copy
    
    With Clone
        
        Set .Left = CloneTree(Node.Left)
        If Not .Left Is Nothing Then _
            Set .Left.Parent = Clone
        
        Set .Right = CloneTree(Node.Right)
        If Not .Right Is Nothing Then _
            Set .Right.Parent = Clone
            
    End With
    
    Set CloneTree = Clone
    
End Function

Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    If Other Is Nothing Then _
        IGeneric_Equals = False: Exit Function
    
    IGeneric_Equals = (Me Is Other)
End Function

Private Function IGeneric_ToString() As String
    With myclass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Elements: " & Me.Elements.Count
    End With
End Function

Private Sub Class_Terminate()
    Call ClearTree(myclass.Root)
    myclass = EmptyStructure
End Sub

'IGenericList
Private Property Get IGenericList_Elements() As IGenericCollection: Set IGenericList_Elements = Me: End Property
Private Function IGenericList_Add(ByVal Element As IGeneric) As Boolean: Call Me.Add(Value:=Element): IGenericList_Add = True: End Function
Private Function IGenericList_AddAll(ByVal Collection As IGenericCollection) As IGenericList: Set IGenericList_AddAll = Me.DoUnionWith(Collection:=Collection): End Function
Private Sub IGenericList_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long, Optional ByVal Index As Long = 0&, Optional ByVal Count As Long = 0&): Call Me.CopyTo(GArray:=GArray, ArrayIndex:=ArrayIndex, Index:=Index, Count:=Count): End Sub
Private Function IGenericList_GetAt(ByVal Index As Long) As IGeneric: Set IGenericList_GetAt = Me.GetAt(Index:=Index): End Function
Private Function IGenericList_IndexOf(ByVal Element As IGeneric) As Long: IGenericList_IndexOf = Me.IndexOf(Value:=Element): End Function

Private Function IGenericList_Pop() As IGeneric
    Call System.Throw(Me, "Private Function IGenericList_Pop").NotImplemented("Not implemented yet")
End Function

Private Function IGenericList_Remove(ByVal Element As IGeneric) As Boolean: IGenericList_Remove = RemoveNode(FindNode(Value:=Element)): End Function
Private Function IGenericList_RemoveAll(ByVal Collection As IGenericCollection) As Long: IGenericList_RemoveAll = Me.DoExeceptWith(Collection:=Collection): End Function
Private Function IGenericList_RetainAll(ByVal Collection As IGenericCollection) As Long: IGenericList_RetainAll = Me.DoIntersectWith(Collection:=Collection): End Function
Private Sub IGenericList_RemoveAt(ByVal Index As Long): Call RemoveNode(FindNodeAt(Rank:=Index)): End Sub

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_Elements() As IGenericCollection: Set IGenericReadOnlyList_Elements = Me: End Property
Private Property Get IGenericReadOnlyList_First() As Long: IGenericReadOnlyList_First = Me.First: End Property
Private Property Get IGenericReadOnlyList_Last() As Long: IGenericReadOnlyList_Last = Me.Last: End Property

Private Function IGenericReadOnlyList_GetAt(ByVal Index As Long) As IGeneric
    
    If (Index >= SizeOf(myclass.Root)) Then _
        Call System.Throw(Me, "Private Function IGenericReadOnlyList_GetAt").IndexOutofRange("(Index > SizeOf(myclass.Root)")

    Set IGenericReadOnlyList_GetAt = FindNodeAt(Rank:=Index).Value
    
End Function

Private Function IGenericReadOnlyList_IndexOf(ByVal Element As IGeneric) As Long: IGenericReadOnlyList_IndexOf = Me.IndexOf(Value:=Element): End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = SizeOf(myclass.Root): End Property

Private Function IGenericCollection_Contains(ByVal Element As IGeneric) As Boolean: IGenericCollection_Contains = Not FindNode(Value:=Element) Is Nothing: End Function

Private Function IGenericCollection_ContainsAll(ByVal Collection As IGenericCollection) As Boolean
    
    Dim Candidate As IGeneric
    Dim Enumerable As IGenericIterator
    
    If (Collection Is Nothing) Then _
        Call System.Throw(Me, "Private Function IGenericCollection_ContainsAll").IsNothing("(Collection Is Nothing)")
    
    If Not (Collection Is Me) Then
                
        If (Collection.Count > 0) Then
            Set Enumerable = Collection.Iterator
        Else
            Set Enumerable = Collection
        End If
        
        With Enumerable
            Do While .HasNext(Candidate)
                If (FindNode(Value:=Candidate) Is Nothing) Then _
                    Exit Function
            Loop
        End With
        
    End If

    IGenericCollection_ContainsAll = True
    
End Function

Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long)
    
    If (GArray Is Nothing) Then _
        Call System.Throw(Me, "Private Sub IGenericCollection_CopyTo").IsNothing("(GArray Is Nothing)")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex >= GArray.Length) Then _
        Call System.Throw(Me, "Private Sub IGenericCollection_CopyTo").IndexOutofRange("(ArrayIndex < GArray.LowerBound) Or (ArrayIndex >= GArray.Length)")
    
    Dim Stack As InternalStack
    Dim Current As GenericNode
    
    Call LoadStack(Stack, myclass.Root, False)

    With GArray
        Do While InOrderTreeWalk(Current, Stack): Set .ElementAt(InterlockedIncrement(ArrayIndex)) = Current.Value: Loop
    End With

End Sub

Private Function IGenericCollection_ToArray() As GenericArray

    Dim Values As GenericArray
    Set Values = GenericArray.Build(SizeOf(myclass.Root))
    
    Call Me.Elements.CopyTo(Values, Values.LowerBound)
    Set IGenericCollection_ToArray = Values

End Function

Private Sub IGenericCollection_Clear(): Call ClearTree(myclass.Root): End Sub
Private Function IGenericCollection_Copy() As IGenericCollection
    
    Dim Clone As Member
    
    Set Clone.Comparer = myclass.Comparer
    Set Clone.Root = CloneTree(myclass.Root)
        
    Set IGenericCollection_Copy = System.CreateInstance(New GenericSortedSet, VarPtr(Clone), LenB(Clone))
    
End Function

Private Function IGenericCollection_Iterator() As IGenericIterator
    
    Dim c As Member

    Call LoadIterator(c.Iterator, myclass.Root, False)
    Set IGenericCollection_Iterator = System.CreateInstance(New GenericSortedSet, VarPtr(c), LenB(c))
    
End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.Current: End Property
Private Function IGenericIterator_HasNext(Optional ByRef Out As IGeneric) As Boolean
    
    Dim Node As GenericNode
    
    If InOrderTreeWalk(Node, myclass.Iterator.Stack) Then
        Set myclass.Iterator.Current = Node.Value
        Set Out = myclass.Iterator.Current
        IGenericIterator_HasNext = True
        Exit Function
    End If
    
    If myclass.Iterator.HasLoaded = False Then _
        Call System.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator

    Set myclass.Iterator.Current = Nothing
    
    IGenericIterator_HasNext = Not Node Is Nothing
    
End Function

Private Sub IGenericIterator_Reset()
    With myclass.Iterator
         If .HasLoaded = False Then _
            Call System.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
          
        Call LoadIterator(myclass.Iterator, .Root, .Stack.Reverse)
    End With
End Sub
