VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSortedSet"
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
'* Class:  GenericSortedSet
'*
'* Purpose: A GenericSortedSet is Red-Black tree based implementation. It provides an efficient means of
'*          storing elements  in sorted order (according to the natural ordering of the values) without affecting performance
'*          as elements are inserted and deleted.
'*          Duplicate Values are not allowed. A Value cannot be null
'*          Changing the sort values of existing items is not supported and may lead to unexpected behavior.
'*          However, if the specified Value already exists in the GenericSortedSet, setting Add(Overwrite:=True) method overwrites the old value.
'*          In contrast, the Add method throws an exception if an element with the specified value already exists.
'*
'* Notes:   A Red Black Tree guarantee that basic operations like search, get, put and remove take logarithmic time O(log n)
'*          Sorted in the ascending order of its values
'*          Implements IGenericReadonlyList
'*
'*
'*======================================================================================================================

'@Folder("<T>Enumerable")

Option Explicit
'@PredeclaredId

Implements IGeneric
Implements IGenericReadOnlyList
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

Private Type InternalStack
    Size As Long
    Items() As GenericNode
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

Public Property Get IsHeightBalanced(): IsHeightBalanced = CheckHeightBalance(myclass.Root): End Property
Public Property Get IsBalanced(): IsBalanced = CheckRedBlackBalance(myclass.Root): End Property
Public Property Get Count() As Long: Count = SizeOf(myclass.Root): End Property
Public Property Get Comparer() As IGenericComparer: Set Comparer = myclass.Comparer: End Property

Public Function Build(Optional ByVal Comparer As IGenericComparer) As GenericSortedSet
    
    If (Me Is GenericSortedSet) = False Then _
        Call System.Throw(Me, "Public Function Build").BuildExeption
    
    If Comparer Is Nothing Then _
        Set Comparer = IGenericComparer
    
    Set myclass.Comparer = Comparer
    Set Build = System.CreateInstance(New GenericSortedSet, VarPtr(myclass), LenB(myclass))
    
End Function

Public Function GetMax() As IGeneric
    Set GetMax = FindMaxOf(myclass.Root)
End Function

Public Function GetMin() As IGeneric
    Set GetMin = FindMinOf(myclass.Root)
End Function

Public Property Get ElementAt(ByVal Index As Long) As IGeneric
Attribute ElementAt.VB_UserMemId = 0

    If (Index < BASE_INDEX) Or (Index > SizeOf(myclass.Root)) Then _
        Call System.Throw(Me, "Public Function ElementAt").IndexOutofRange("(Index < BASE_INDEX) Or (Index > SizeOf(myclass.Root))")

    Set ElementAt = SelectKth(myclass.Root, Index - BASE_INDEX).Value
    
End Property

Public Function GetRankOf(ByVal Value As IGeneric) As Long
    GetRankOf = RankOf(Value, myclass.Root) + BASE_INDEX
End Function

Public Function IndexOf(ByVal Value As IGeneric) As Long
    
    Dim Current As GenericNode: Set Current = myclass.Root
    Dim Order As CompareResult
    
    Do While Current Is Nothing = False
        Order = myclass.Comparer.Compare(Value, Current.Value)
        
        If (Order = 0) Then _
            IndexOf = IndexOf + SizeOf(Current.Left) + BASE_INDEX: Exit Function
            
        If Order < 0 Then
            Set Current = Current.Left
        Else
            IndexOf = IndexOf + 1 + SizeOf(Current.Left): Set Current = Current.Right
        End If
    Loop
   
    IndexOf = -1
        
End Function

Private Function RankOf(ByVal Value As IGeneric, ByVal Node As GenericNode) As Long
    
    Dim Order As CompareResult
    
    Do While Node Is Nothing = False
        Order = myclass.Comparer.Compare(Value, Node.Value)
        
        If (Order = 0) Then _
            RankOf = RankOf + SizeOf(Node.Left): Exit Function
            
        If Order < 0 Then
            Set Node = Node.Left
        Else
            RankOf = RankOf + 1 + SizeOf(Node.Left): Set Node = Node.Right
        End If
    Loop
    
'    If Node Is Nothing Then _
'        Exit Function
'
'    Select Case myclass.Comparer.Compare(Value, Node.Value)
'        Case Is < 0
'            RankOf = RankOf(Value, Node.Left)
'        Case Is > 0
'            RankOf = 1 + SizeOf(Node.Left) + RankOf(Value, Node.Right)
'        Case Else
'            RankOf = SizeOf(Node.Left)
'    End Select

End Function

Public Function GetHigherThan(ByVal Value As IGeneric) As IGeneric
    
    Dim Current As GenericNode: Set Current = myclass.Root
    Dim BestSoFar As GenericNode
    
    With myclass.Comparer
        Do While (Current Is Nothing) = False
            If .Compare(Value, Current.Value) = IsLower Then
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
            If .Compare(Value, Current.Value) = IsGreater Then
                Set BestSoFar = Current
                Set Current = Current.Right
            Else
                Set Current = Current.Left
            End If
        Loop
    End With
    
    Set GetLowerThan = BestSoFar.Value
    
End Function

Private Function SelectKth(ByVal Node As GenericNode, ByVal Rank As Long) As GenericNode
    
    Dim leftSize As Long
    
    If (Node Is Nothing) Then _
        Set SelectKth = Nothing: Exit Function

    If Node.Left Is Nothing = False Then _
        leftSize = Node.Left.Count
    
    Select Case True
        Case leftSize > Rank
            Set SelectKth = SelectKth(Node.Left, Rank)
        Case leftSize < Rank
            Set SelectKth = SelectKth(Node.Right, Rank - leftSize - 1)
        Case Else
            Set SelectKth = Node
    End Select
    
End Function

Private Function FindMaxOf(ByVal Node As GenericNode) As GenericNode
    
    If Not (Node Is Nothing) Then
        Do While (Node.Right Is Nothing) = False
            Set Node = Node.Right
        Loop
    End If
    
    Set FindMaxOf = Node

End Function

Private Function FindMinOf(ByVal Node As GenericNode) As GenericNode
    
    If Not (Node Is Nothing) Then
        Do While (Node.Left Is Nothing) = False
            Set Node = Node.Left
        Loop
    End If
    
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

Public Sub AddAll(ByVal Sequence As IGenericIterator)
    
    Dim Element As IGeneric

    With Sequence
        Do While .HasNext(Element): Call Me.Add(Element, True): Loop
    End With

End Sub

Public Function ToArray(Optional ByVal Reverse As Boolean) As GenericArray
   
    Dim Values As GenericArray
    Set Values = GenericArray.Build(Me.Count)
    
    Dim Index As Long: Index = Values.LowerBound - 1
    
    Dim Stack As InternalStack
    Dim Current As GenericNode

    Call LoadStack(Stack, myclass.Root, Reverse)

    Do While InOrderTreeWalk(Current, Stack): Set Values.ElementAt(InterlockedIncrement(Index)) = Current.Value: Loop
    
    Set ToArray = Values
    
End Function

Private Sub AddValuesToArray(ByVal Node As GenericNode, ByVal GArray As GenericArray, ByRef Index As Long)

    If Not (Node.Left Is Nothing) Then _
        Call AddValuesToArray(Node.Left, GArray, Index)
    
    Set GArray(Index) = Node.Value: Index = Index + 1
    
    If Not (Node.Right Is Nothing) Then _
        Call AddValuesToArray(Node.Right, GArray, Index)
    
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
        
    lh = SizeOf(Node.Left)
    rh = SizeOf(Node.Right)
    
    'High balanced
    If Abs(lh - rh) <= 1 Then _
        If CheckHeightBalance(Node.Left) And CheckHeightBalance(Node.Left) Then _
            CheckHeightBalance = True
  
End Function

Public Function Iterator(Optional ByVal Reverse As Boolean) As IGenericIterator
    
    Dim c As Member

    Call LoadIterator(c.Iterator, myclass.Root, Reverse)
    Set Iterator = System.CreateInstance(New GenericSortedSet, VarPtr(c), LenB(c))
    
End Function

Public Sub Clear()
    
    If myclass.Root Is Nothing = False Then
        With myclass.Root
            If .Left Is Nothing = False Then _
                Call .Left.Clear
            Set .Left = Nothing
    
            If .Right Is Nothing = False Then _
                Call .Right.Clear
            Set .Right = Nothing
        End With
    End If
    
End Sub

Public Sub IntersectWith(ByVal Collection As IGenericCollection)
    
    Dim Value As IGeneric
    Dim NodesToSave As GenericList
    Dim Node As GenericNode
    Dim Current As GenericNode
    Dim Parent  As GenericNode
    Dim GrandParent As GenericNode
    Dim GreatGrandParent As GenericNode
    Dim Order As Long
    
    If Collection Is Nothing Then
    'Error
    End If
    
    If Collection.Count = 0 Then
        Exit Sub
    End If
    
    If IGeneric_IsRelatedTo(Collection) Then
        'Special handling:exit sub
    End If
    
    Set NodesToSave = GenericList.Build(Me.Count)
    
    With Collection.Iterator
    
        Do While .HasNext(Value)
            Set Node = FindNode(Value)
            If Node Is Nothing = False Then
                Call Me.Remove(Value)
                With Node 'reset
                    '  Set .Value = Value 'not necessary -if not reference equal?!
                    Set .Parent = Nothing
                    Set .Left = Nothing
                    Set .Right = Nothing
                    .Count = 1
                    .IsRed = False
                End With
                Call NodesToSave.Add(Node)
            End If
        Loop
        
    End With
    
    Call Me.Clear
    
    With NodesToSave.Iterator
    
        If .HasNext(Node) Then _
            Set myclass.Root = Node
        
        Do While .HasNext(Node)
            Set Current = myclass.Root
            Set Value = Node.Value
            
            With myclass.Comparer
            
                Do
                    Order = .Compare(Value, Current.Value)
        
                    Current.Count = Current.Count + 1
                    
                    If Split4Node(Current) Then _
                        If (Parent Is Nothing) = False Then _
                            If Parent.IsRed Then _
                                Call EnsureBalance(Current, Parent, GrandParent, GreatGrandParent)
                    
                    Set GreatGrandParent = GrandParent
                    Set GrandParent = Parent
                    Set Parent = Current
                    
                    If Order < 0 Then
                        Set Current = Current.Left
                    Else
                        Set Current = Current.Right
                    End If
            
                Loop While (Current Is Nothing) = False
                
            End With
        '    If (Parent Is Nothing) Then _
        '        Call System.Throw(Me, "Private Sub InsertionBalance").NullPointer("Parent cannot be null here!")
            
            If (Order < 0) Then
                Set Parent.Left = Node
            Else
                Set Parent.Right = Node
            End If
            
            Set Node.Parent = Parent
            Node.IsRed = True
            
            If (Parent.IsRed) Then _
                Call EnsureBalance(Node, Parent, GrandParent, GreatGrandParent)
            
            myclass.Root.IsRed = False
        Loop
        
    End With
    
End Sub

Public Function Remove(ByVal Value As IGeneric) As Boolean


End Function

Public Function Contains(ByVal Value As IGeneric) As Boolean: Contains = Not FindNode(Value:=Value) Is Nothing: End Function

Private Function FindNode(ByVal Value As IGeneric) As GenericNode
    
    Dim Current As GenericNode: Set Current = myclass.Root
    Dim Order As Long
    
    If (Value Is Nothing) Then _
        Call System.Throw(Me, "Private Function FindNode").IsNothing("(Value Is Nothing)")
    
    With myclass.Comparer
        Do While (Current Is Nothing) = False
        
            Order = .Compare(Value, Current.Value)
           
            If (Order = 0) Then _
                Set FindNode = Current: Exit Function
            
            If Order < 0 Then
                Set Current = Current.Left
            Else
                Set Current = Current.Right
            End If
            
        Loop
    End With
    
End Function

Private Function MaxHeightOf(ByVal Root As GenericNode) As Long
    MaxHeightOf = 2 * LOGn(SizeOf(Root) + 1, Base:=2)
End Function
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
'leftleaning
Private Function InsertLeftLeaning(ByVal Node As GenericNode, ByVal Value As IGeneric) As GenericNode
    
    If Node Is Nothing Then _
        Set InsertLeftLeaning = GenericNode.Build(Value:=Value, IsRed:=True): Exit Function
    
    With Node
        .Count = .Count + 1
        Select Case myclass.Comparer.Compare(Value, .Value)
            Case Is < 0: Set .Left = InsertLeftLeaning(.Left, Value)
            Case Is > 0: Set .Right = InsertLeftLeaning(.Right, Value)
            Case Else: Set .Value = Value
        End Select
    End With
    
    If (IsRed(Node.Right)) Then _
        Set Node = RotateLeft(Node): Node.IsRed = Node.Left.IsRed: Node.Left.IsRed = True
    
    If (IsRed(Node.Left)) Then _
        If IsRed(Node.Left.Left) Then _
            Set Node = RotateRight(Node): Node.IsRed = Node.Right.IsRed: Node.Right.IsRed = True
    
    With Node
        If (IsRed(.Left)) Then _
            If IsRed(.Right) Then _
                .IsRed = Not .IsRed: .Left.IsRed = False: .Right.IsRed = False
                'Call FlipColor(Node)
    End With
    
    Set InsertLeftLeaning = Node
    
End Function

Public Sub AddLeftLeaning(ByVal Value As IGeneric)
    Set myclass.Root = InsertLeftLeaning(myclass.Root, Value)
    myclass.Root.IsRed = False
End Sub

Public Function Add( _
                        ByVal Value As IGeneric, _
                        Optional ByVal Overwrite As Boolean) As Boolean
    
    Dim Node As GenericNode
    Dim Current As GenericNode
    Dim Parent  As GenericNode
    Dim GrandParent As GenericNode
    Dim GreatGrandParent As GenericNode
    Dim Order As Long, i As Long
    
    If (Value Is Nothing) Then _
        Call System.Throw(Me, "Public Function Add").IsNothing("(Value Is Nothing)")
    
    If (myclass.Root Is Nothing) Then _
        Set myclass.Root = GenericNode.Build(Value:=Value, Parent:=Nothing, IsRed:=False): _
        Add = True: _
        Exit Function

    Set Current = myclass.Root
    With myclass.Comparer
        Do
            Order = .Compare(Value, Current.Value)
           
            If (Order = 0) Then
                myclass.Root.IsRed = False
                Add = False
                Do While Parent Is Nothing = False
                    Parent.Count = Parent.Count - 1
                    Set Parent = Parent.Parent
                Loop
                
                If (Overwrite) = False Then _
                    Call System.Throw(Me, "Public Function Add").DuplicateNotAllowed(Value)
        
                Exit Function
            End If
            
            Current.Count = Current.Count + 1
            
            If Split4Node(Current) Then _
                If (Parent Is Nothing) = False Then _
                    If Parent.IsRed Then _
                        Call EnsureBalance(Current, Parent, GrandParent, GreatGrandParent)
            
            Set GreatGrandParent = GrandParent
            Set GrandParent = Parent
            Set Parent = Current
            
            If Order < 0 Then
                Set Current = Current.Left
            Else
                Set Current = Current.Right
            End If
    
        Loop While (Current Is Nothing) = False
    End With
'    If (Parent Is Nothing) Then _
'        Call System.Throw(Me, "Private Sub InsertionBalance").NullPointer("Parent cannot be null here!")

    Set Node = GenericNode.Build(Value:=Value, Parent:=Parent, IsRed:=True)
    
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

Private Sub UpdateSizeOf(ByVal Node As GenericNode)

    Dim i As Long, j As Long
    
    With Node
        If .Left Is Nothing = False Then _
            i = .Left.Count
        
        If .Right Is Nothing = False Then _
            j = .Right.Count
        
        .Count = 1 + i + j
    End With

End Sub

Private Function Is4Node(ByVal Node As GenericNode) As Boolean
    If (IsRed(Node.Left)) Then _
        Is4Node = IsRed(Node.Right)
End Function

Private Function IsRed(ByVal Node As GenericNode) As Boolean
    If Not (Node Is Nothing) Then _
        IsRed = Node.IsRed
End Function

Private Function IsBlack(ByVal Node As GenericNode) As Boolean
    If Not (Node Is Nothing) Then _
        IsBlack = Not Node.IsRed
End Function

Private Function IsNothingOrBlack(ByVal Node As GenericNode) As Boolean
    If (Node Is Nothing) Then _
        IsNothingOrBlack = True: Exit Function
    
    IsNothingOrBlack = Not Node.IsRed
End Function

Private Function Split4Node(ByVal Node As GenericNode) As Boolean
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

Private Sub FlipColor(ByVal Node As GenericNode)
    
    If Not Node Is Nothing Then
        With Node
            .IsRed = Not .IsRed
            .Left.IsRed = Not .Left.IsRed
            .Right.IsRed = Not .Right.IsRed
        End With
    End If

End Sub

Private Function RotateLeft(ByVal x As GenericNode) As GenericNode

    Dim y  As GenericNode
    
    With x
        Set y = .Right
        Set .Right = y.Left
        
        If Not y.Left Is Nothing Then _
            Set y.Left.Parent = x
        
        Set y.Parent = .Parent
        Set y.Left = x: Set .Parent = y
        
        y.Count = .Count: Call UpdateSizeOf(x)
    End With
    
    'left leaning
'    With y
'        .IsRed = .Left.IsRed
'        .Left.IsRed = True
'    End With
    
    Set RotateLeft = y
    
End Function

Private Function RotateRight(ByVal x As GenericNode) As GenericNode

    Dim y  As GenericNode
    
    With x
        Set y = .Left
        Set .Left = y.Right
        
        If Not y.Right Is Nothing Then _
            Set y.Right.Parent = x
        
        Set y.Parent = .Parent
        Set y.Right = x: Set .Parent = y
        
        y.Count = .Count: Call UpdateSizeOf(x)
    End With
     'left leaning
'    With y
'        .IsRed = .Right.IsRed
'        .Right.IsRed = True
'    End With
'
    Set RotateRight = y
    
End Function

Private Function RotateLeftRight(ByVal Node As GenericNode) As GenericNode
    
    Set Node.Left = RotateLeft(Node.Left)
    Set RotateLeftRight = RotateRight(Node)

End Function

Private Function RotateRightLeft(ByVal Node As GenericNode) As GenericNode

    Set Node.Right = RotateRight(Node.Right)
    Set RotateRightLeft = RotateLeft(Node)

End Function

Private Function InOrderTreeWalk( _
                                    ByRef Current As GenericNode, _
                                    ByRef Stack As InternalStack) As Boolean
    
    Dim Node As GenericNode

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

Private Sub LoadStack(ByRef Stack As InternalStack, ByVal Node As GenericNode, ByVal Reverse As Boolean)
    
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

Private Sub LoadIterator(ByRef e As Iterator, ByVal Root As GenericNode, ByVal Reverse As Boolean)
    
    If (Root Is Nothing) Then _
        Call System.Throw(Me, "Private Sub LoadIterator").IsNothing("(Root Is Nothing)")
    
    With e
        Set .Current = Nothing
        Set .Root = Root
        Call LoadStack(.Stack, .Root, Reverse)
        .HasLoaded = True
    End With
    
End Sub

Private Function CopyOf(ByVal GArray As GenericArray, ByVal ArrayIndex As Long, ByVal Caller As String) As GenericArray
    
    If (GArray Is Nothing) Then _
        Call System.Throw(Me, Caller).IsNothing("(GArray Is Nothing)")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call System.Throw(Me, Caller).IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < SizeOf(myclass.Root)) Then _
        Call System.Throw(Me, Caller).InvalidInput("(GArray.Length - Index < myClass.Count)")
    
    Dim Stack As InternalStack
    Dim Node As GenericNode
    Dim Current As IGeneric
    
    Call LoadStack(Stack, myclass.Root, False)

    Do While InOrderTreeWalk(Node, Stack)
        Set GArray(ArrayIndex) = Node.Value
        ArrayIndex = ArrayIndex + 1
    Loop
       
    Set CopyOf = GArray
    
End Function

'IGeneric
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GenericSortedSet: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericSortedSet):  End Property
Private Function IGeneric_Clone() As IGeneric
    
    Dim Clone As Member
    Clone = myclass
    Set Clone.Root = CloneTree(myclass.Root)
        
    Set IGeneric_Clone = System.CreateInstance(New GenericSortedSet, VarPtr(Clone), LenB(Clone))
    
End Function

Private Function CloneTree(ByVal Node As GenericNode) As GenericNode
    
    Dim newNode As GenericNode
    
    If Node Is Nothing Then _
        Exit Function
    
    With Node
        Set newNode = GenericNode.Build(Value:=.Value, Parent:=.Parent, IsRed:=.IsRed, Count:=.Count)
        Set newNode.Left = CloneTree(.Left)
        Set newNode.Right = CloneTree(.Right)
    End With
    
    Set CloneTree = newNode
    
End Function

Private Sub IGeneric_Dispose()
    Call Me.Clear
End Sub
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
                            "Elements: " & Me.Count
    End With
End Function

Private Sub Class_Terminate()

    If Not myclass.Root Is Nothing Then _
        Call System.Dispose(myclass.Root)
    
    myclass = EmptyStructure
    
End Sub

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_BaseIndex() As Long: IGenericReadOnlyList_BaseIndex = BASE_INDEX: End Property
Private Property Get IGenericReadOnlyList_Count() As Long: IGenericReadOnlyList_Count = Me.Count: End Property
Private Property Get IGenericReadOnlyList_ElementAt(ByVal Index As Long) As IGeneric: Set IGenericReadOnlyList_ElementAt = Me.ElementAt(Index): End Property
Private Sub IGenericReadOnlyList_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="IGenericReadOnlyList_CopyTo"):   End Sub
Private Function IGenericReadOnlyList_IndexOf(ByVal Item As IGeneric) As Long: IGenericReadOnlyList_IndexOf = Me.IndexOf(Value:=Item): End Function
Private Function IGenericReadOnlyList_Iterator() As IGenericIterator: Set IGenericReadOnlyList_Iterator = Me.Iterator(Reverse:=False): End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.Contains(Value:=Item): End Function
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = Me.ToArray: End Function
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="Private Sub IGenericCollection_CopyTo"): End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Iterator(Reverse:=False): End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.Current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    Dim Node As GenericNode
    
    With myclass.Iterator
        If InOrderTreeWalk(Node, .Stack) Then
            Set .Current = Node.Value
            Set Out = .Current
            IGenericIterator_HasNext = True
            Exit Function
        End If
        
        If .HasLoaded = False Then _
            Call System.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator
        
        Set Out = Nothing
        Set .Current = Nothing
      
    End With
 
End Function

Private Sub IGenericIterator_Reset()
    With myclass.Iterator
         If .HasLoaded = False Then _
            Call System.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
          
        Call LoadIterator(myclass.Iterator, .Root, .Stack.Reverse)
    End With
End Sub


