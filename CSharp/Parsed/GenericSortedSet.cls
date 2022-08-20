VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSortedSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@IgnoreModule ProcedureNotUsed, AssignedByValParameter, ArgumentWithIncompatibleObjectType, UnassignedVariableUsage, SuspiciousPredeclaredInstanceAccess, ProcedureCanBeWrittenAsFunction
'@Folder "Collections"
Option Explicit

'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericSortedSet
'*
'*
'* Purpose: A GenericSortedSet class is based on the model of mathematical sets and provides high-performance set operations, such as set addition (unions) and set subtraction (except).
'*          A set is a collection that contains no duplicate elements. In simple terms, the GenericSortedSet class can be thought of as a sorted map collection without values.
'*
'*          This implementation differs from a ordinary set (pseudo random order determined by hashes calculated from the elements)
'*          in that it maintains a self-balancing binary search tree, more specifically a Red-Black tree. It provides an efficient means of
'*          storing elements  in sorted Order (according to the natural ordering of the elements) without affecting performance as elements are inserted and deleted.
'*          Duplicate elements are not allowed. If a element can be null,depends on the current comparer.
'*          Changing the sort values of existing Elements is not supported and may lead to unexpected behavior.
'*          If the specified element already exists in the GenericSortedSet, the TryAdd method will not throw an exception.
'*          We can quickly find the ith smallest number in a set or the rank of a given element in the total ordering of the set.
'*
'*          RB-trees are binary trees satisfying some additional constraints:
'*              1. Every node is colored either red or black. (Equivalently, the edge from the node to its parent is colored red or black.)
'*              2. Leaf nodes (null pointers) are black.
'*              3. No consecutive reds - if a node is red, its parent is black.
'*              4. Every leaf (null) has the same "black depth" - the number of black nodes on the path from the leaf to the root.
'*
'* Notes:  The implementation is based on a Red Black Tree -> guarantee that basic operations like search, get, put and remove take logarithmic time O(log n)
'*         In a set, the elements must be unique -> duplicate elements are not allowed.
'*         Elements must be immutable as long as they are used in the GenericSortedSet.
'*         It provides many mathematical set operations, such as intersection, union, and (symmetric) difference (except).
'*         It maintains (ascending) Order (according to the associated Comparer implementation) and does not store duplicate elements.
'*         TryAdd (TryRemove) method returned a boolean that indicates if it was successfully added (removed).
'*         This tree implementation supports general Order-statistic operations.
'*         Every node knows the number of items in its subtree. This permits O(log n) implementations of search-by-index,
'*         IndexOf(x), and other index-based operations.
'*
'*         If NULL values are expected in the collection, a NULL comparison operator must be provided. Otherwise null pointer exceptions can occur!
'*
'*
'*======================================================================================================================

Implements IGeneric
Implements IGenericCollection
Implements IGenericSequence
Implements IGenericIterator
Implements IGenericPredicator

#If Win64 Then
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

Private Const LOG2 As Double = 0.693147180559945  'Dont change!

Private Const BASE_INDEX As Long = 0
'@Ignore IntegerDataType
Private Const SORT_ORDER As Integer = SortOrder.Ascending 'Ascending
Private Const IS_BLACK As Byte = 0
Private Const IS_RED As Byte = 2 ^ 0
Private Const IS_MARKED As Byte = 2 ^ 1

Private Type InternalStack
    Size As Long
    Elements() As GenericTreeNode
    Reverse As Boolean
End Type

Private Type Iterator
    SortedSet As GenericSortedSet
    Current As Object
    Stack As InternalStack
    HasLoaded As Boolean
End Type

Private Type Member
    Root As GenericTreeNode
    Comparer As IGenericComparer
    Iterator As Iterator
End Type
Private This As Member

Public Property Get Instance() As IGeneric: Set Instance = Me: End Property
Public Property Get Elements() As IGenericCollection: Set Elements = Me: End Property
Public Property Get First() As Long: First = BASE_INDEX: End Property

Public Property Get Last() As Long
    
    Dim c As Long
    Last = BASE_INDEX
    c = SizeOf(This.Root)
    
    If c > 0 Then _
        Last = c - Last - 1
   
End Property

Public Property Get Direction() As SortOrder: Direction = SORT_ORDER: End Property
Public Property Get IsHeightBalanced() As Boolean: IsHeightBalanced = CheckHeightBalance(This.Root): End Property
Public Property Get IsBalanced() As Boolean: IsBalanced = CheckRedBlackBalance(This.Root): End Property
Public Property Get Comparer() As IGenericComparer: Set Comparer = This.Comparer: End Property
Public Property Get SortedSetEquality() As IGenericPredicator: Set SortedSetEquality = GenericSortedSet: End Property

Public Function Build(Optional ByVal Comparer As IGenericComparer) As GenericSortedSet
        
    Dim newSet As Member
    
    Set newSet.Comparer = Comparer

    Set Build = New GenericSortedSet
    Call System.Inject(Build, VarPtr(newSet), LenB(newSet))
    
End Function

Public Function BuildFrom(ByVal Sequence As IGenericSequence, Optional ByVal Comparer As IGenericComparer) As GenericSortedSet
    
    Dim Other As GenericSortedSet
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function BuildFrom").IsNothing("Sequence")
        
    If TypeOf Sequence Is GenericSortedSet Then _
        If IsComparable(Me, Sequence) Then _
            Set Other = Sequence
            
    If (Other Is Nothing) Then
        Set BuildFrom = GenericSortedSet.Build(Comparer:=Comparer).DoUnion(Sequence)
    Else
        Set BuildFrom = Other.Elements.Copy
    End If
    
End Function

Public Function Of(ByVal Comparer As IGenericComparer, ParamArray IGenericValues() As Variant) As IGenericCollection
        
    Dim SortedSet As GenericSortedSet
    Dim i As Long
    
    If IsMissing(IGenericValues) Then _
        Call GenericError.Build(Me, "Public Function Of").NoSuchElement("ParamArray IGenericValues are missing")
    
    Set SortedSet = GenericSortedSet.Build(Comparer:=Comparer)
    
    With SortedSet
        For i = LBound(IGenericValues) To UBound(IGenericValues)
            
            If Not VBA.IsObject(IGenericValues(i)) Then _
                Call GenericError.Build(Me, "Public Function Of").InvalidInput("Datatype {" & TypeName$(IGenericValues(i)) & "} not allowed")
            
            If (IGenericValues(i) Is Nothing) Then _
                Call GenericError.Build(Me, "Public Function Of").InvalidInput("Element is nothing")
            
            If Not TypeOf IGenericValues(i) Is IGeneric Then _
                Call GenericError.Build(Me, "Public Function Of").InvalidInput("Not TypeOf IGenericValues(i) Is IGeneric")
                
            Call .TryAdd(Element:=IGenericValues(i))
            
        Next
    End With
    
    Set Of = SortedSet
    
End Function

Public Function GetMax() As Object
   
    If Not This.Root Is Nothing Then _
        Set GetMax = This.Root.GetMax.Element

End Function

Public Function GetMin() As Object
    
    If Not This.Root Is Nothing Then _
        Set GetMin = This.Root.GetMin.Element

End Function

Public Function GetAt(ByVal Index As Long) As Object
    
    If (This.Root Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function GetAt").IndexOutofRange("(Index >= Count")
        
    If (Index >= This.Root.Count) Then _
        Call GenericError.Build(Me, "Public Function GetAt").IndexOutofRange("(Index >= Count")
    
    Set GetAt = This.Root.GetAt(Index).Element
    
End Function

Public Function GetRankOf(ByVal Element As Object) As Long
    
    If Not This.Root Is Nothing Then
        GetRankOf = This.Root.GetRankOf(Element, This.Comparer)
    Else
        GetRankOf = -1
    End If
    
End Function

Public Function GetIndexOf(ByVal Element As Object) As Long
    
    If Not This.Root Is Nothing Then
        GetIndexOf = This.Root.GetIndexOf(Element, This.Comparer)
    Else
        GetIndexOf = -1
    End If
        
End Function

Public Function GetAfter(ByVal Element As Object) As Object
    
    If Not This.Root Is Nothing Then _
        Set GetAfter = This.Root.GetAfter(Element, This.Comparer)
    
End Function

Public Function GetBefore(ByVal Element As Object) As Object
   
    If Not This.Root Is Nothing Then _
        Set GetBefore = This.Root.GetBefore(Element, This.Comparer)
    
End Function

Public Sub CopyTo( _
                    ByVal Target As GenericArray, _
                    ByVal Index As Long, _
                    Optional ByVal BeginAt As Long = BASE_INDEX, _
                    Optional ByVal Count As Long = 0)
    
    If (Target Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function CopyTo").IsNothing("Target")
    
    If (Target.IsEmpty) Then _
        Call GenericError.Build(Me, "Public Function CopyTo").IsNothing("(Target.IsEmpty)")
        
    If (This.Root Is Nothing) Then _
        Exit Sub
        
    Call This.Root.CopyTo(Target:=Target, Index:=Index, BeginAt:=BeginAt, Elements:=Count)
    
End Sub

Public Function GetLessThan(ByVal Max As Object) As GenericSortedSet

    Dim Current As GenericTreeNode
    
    Dim SubSet As GenericSortedSet
    Set SubSet = GenericSortedSet.Build(Comparer:=This.Comparer)
    
    If Not This.Root Is Nothing Then
        
        Set Current = This.Root.GetAfter(Max, This.Comparer).Before
        With SubSet
            Do
                Call .TryAdd(Current.Element)
                Set Current = Current.Before
            Loop While Not (Current Is Nothing)
        End With
        
    End If
    
    Set GetLessThan = SubSet
    
End Function

Public Function GetGreaterThan(ByVal Min As Object) As GenericSortedSet

    Dim Current As GenericTreeNode
    
    Dim SubSet As GenericSortedSet
    Set SubSet = GenericSortedSet.Build(Comparer:=This.Comparer)
    
    If Not This.Root Is Nothing Then
        
        Set Current = This.Root.GetBefore(Min, This.Comparer).After
        With SubSet
            Do
                Call .TryAdd(Current.Element)
                Set Current = Current.After
            Loop While Not (Current Is Nothing)
        End With
        
    End If
    
    Set GetGreaterThan = SubSet
    
End Function
'TODO Check loop
Public Function GetBetween(ByVal Min As Object, ByVal Max As Object) As GenericSortedSet
   
    Dim Value As IGenericValue
    Dim Current As GenericTreeNode
    
    Dim SubSet As GenericSortedSet
   
    If Not This.Root Is Nothing Then
        If (This.Comparer Is Nothing) Then
            
            Set Value = Min 'try cast
            If Not Value.IsGreaterThan(This.Root.GetMin.Element) Then
                Set GetBetween = Me.GetLessThan(Max)
                Exit Function
            End If
        
            Set Value = Max 'try cast
            If Not Value.IsLessThan(This.Root.GetMax.Element) Then
                Set GetBetween = Me.GetGreaterThan(Min)
                Exit Function
            End If
            
            Set SubSet = GenericSortedSet.Build()
            
            Set Current = This.Root.GetBefore(Min).After
            
            With Value
                Do
                    If .IsLessThan(Current.Element) Then _
                        Exit Do
                    Call SubSet.TryAdd(Current.Element)
                    Set Current = Current.After
                Loop While Not (Current Is Nothing)
            End With
            
        Else
            
            If Not This.Comparer.AscendingOrdered(Max, This.Root.GetMax.Element) Then
                Set GetBetween = Me.GetGreaterThan(Min)
                Exit Function
            End If
            
            If Not This.Comparer.AscendingOrdered(This.Root.GetMin.Element, Min) Then
                Set GetBetween = Me.GetLessThan(Max)
                Exit Function
            End If
        
            Set SubSet = GenericSortedSet.Build(Comparer:=This.Comparer)
            
            Set Current = This.Root.GetBefore(Min, This.Comparer).After
            
            With This.Comparer
                Do
                    If .AscendingOrdered(Max, Current.Element) Then _
                        Exit Do
                    Call SubSet.TryAdd(Current.Element)
                    Set Current = Current.After
                Loop While Not (Current Is Nothing)
            End With
            
        End If
    End If
     
    Set GetBetween = SubSet
    
End Function

Public Function Equals(ByVal Other As GenericSortedSet) As Boolean: Equals = GenericSortedSet.SortedSetEquality.IsTrue(Me, Other): End Function

' Bulk Set Operations
'Todo Bottom up implementaton
Public Function DoUnion(ByVal Sequence As IGenericSequence) As GenericSortedSet

    '@Ignore VariableNotAssigned
    Dim Candidate As Object
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function DoUnion").IsNothing("Sequence")
    
    Set DoUnion = Me
    
    If (Sequence Is Me) Then _
        Exit Function
            
    With Sequence.Iterator
        Do While .HasNext(Candidate): Call Me.TryAdd(Candidate): Loop
    End With
    
End Function

Public Function DoExcept(ByVal Sequence As IGenericSequence) As GenericSortedSet
    
    Call GenericError.Build(Me, "Public Function DoExecept").NotImplemented
    
    Dim Collection As IGenericCollection, Other As GenericSortedSet
    '@Ignore VariableNotAssigned
    '@Ignore VariableNotAssigned
    Dim Candidate As Object, Element As IGenericValue
    Dim Min As Object
    Dim Max As Object
 
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function DoExecept").IsNothing("Sequence")
    
    Set DoExcept = Me
    
    If (This.Root Is Nothing) Then _
        Exit Function
    
    If (Sequence Is Me) Then _
        Call Me.Elements.Clear: Exit Function
    
    If TypeOf Sequence Is IGenericCollection Then
        '@Ignore SetAssignmentWithIncompatibleObjectType
        Set Collection = Sequence
        
        If (Collection.Count = 0) Then _
            Exit Function
            
    End If
    
    Set Min = Me.GetMin
    Set Max = Me.GetMax
    
    If TypeOf Sequence Is GenericSortedSet Then
        If IsComparable(Me, Sequence) Then
            
            Set Other = Sequence
            If (This.Comparer Is Nothing) Then
            
                If Comparable(Min).IsGreaterThan(Other.GetMax) Then _
                    Exit Function
        
                If Comparable(Max).IsLessThan(Other.GetMin) Then _
                    Exit Function
            Else
        
                If This.Comparer.AscendingOrdered(Other.GetMax, Min) Then _
                    Exit Function
        
                If This.Comparer.AscendingOrdered(Max, Other.GetMin) Then _
                    Exit Function
        
            End If
            
        End If
    End If
      
    With Sequence.Iterator
        If (This.Comparer Is Nothing) Then
            
            Do While .HasNext(Element) 'Try cast
                If Not (Element Is Nothing) Then _
                    If Not Element.IsLessThan(Min) Then _
                        If Not Element.IsGreaterThan(Max) Then _
                            Call RemoveNode(This.Root.Find(Element))
            Loop
        
        Else
        
            Do While .HasNext(Candidate)
                If Not This.Comparer.AscendingOrdered(Candidate, Min) Then _
                    If Not This.Comparer.AscendingOrdered(Max, Candidate) Then _
                        Call RemoveNode(This.Root.FindByComparer(Candidate, This.Comparer))
                        
            Loop
           
        End If
    End With
   
End Function

Public Function DoSymmetricExcept(ByVal Sequence As IGenericSequence) As GenericSortedSet
    
    Call GenericError.Build(Me, "Public Function DoSymmetricExcept").NotImplemented
    
    '@Ignore VariableNotAssigned
    Dim Candidate As Object
    Dim Collection As IGenericCollection, Other As GenericSortedSet
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function DoSymmetricExcept").IsNothing("Sequence")
    
    Set DoSymmetricExcept = Me
    
    If (Sequence Is Me) Then _
        Call Me.Elements.Clear: Exit Function
    
    If (This.Root Is Nothing) Then _
        Call Me.DoUnion(Sequence): Exit Function

    If TypeOf Sequence Is IGenericCollection Then
        '@Ignore SetAssignmentWithIncompatibleObjectType
        Set Collection = Sequence
        
        If (Collection.Count = 0) Then _
            Exit Function
         
    End If
    
    If TypeOf Sequence Is GenericSortedSet Then _
        If IsComparable(Me, Sequence) Then _
            Set Other = Sequence
                
    If (Other Is Nothing) Then
    
        Call GenericError.Build(Me, "Public Function DoSymmetricExcept").NotImplemented("Not (ComparerEquality(This.Comparer, Other.Comparer)")
    
    Else
    
        With Sequence.Iterator
            If Not .HasNext(Candidate) Then _
                Exit Function
                
            If (This.Comparer Is Nothing) Then
            
                Do
                
                    If Not RemoveNode(This.Root.Find(Candidate)) Then _
                        Call Me.TryAdd(Candidate)
                        
                Loop While .HasNext(Candidate)
                
            Else
            
                Do
                
                    If Not RemoveNode(This.Root.FindByComparer(Candidate, This.Comparer)) Then _
                        Call Me.TryAdd(Candidate)
                        
                Loop While .HasNext(Candidate)
            
            End If
        End With
    
    End If
   
End Function

Public Function DoIntersect(ByVal Sequence As IGenericSequence) As GenericSortedSet
    
    Dim Collection As IGenericCollection, Other As GenericSortedSet
    Dim MaxSize As Long
    
    Dim OtherIterator As IGenericIterator
    Dim thisEnded As Boolean, otherEnded As Boolean
    Dim Max As Object
        
    Dim SavedNodes As InternalStack, Stack As InternalStack
    Dim SavedNode As GenericTreeNode, Current As GenericTreeNode
    
    Dim Parent  As GenericTreeNode, GrandParent As GenericTreeNode, GreatGrandParent As GenericTreeNode
    Dim IsLower As Boolean
    
    Dim Element As Object
    Dim Value As IGenericValue
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function DoIntersect").IsNothing("Sequence")
    
    Set DoIntersect = Me
    
    If (This.Root Is Nothing) Then _
        Exit Function
    
    If (Sequence Is Me) Then _
        Exit Function
    
    MaxSize = This.Root.Count
    
    If TypeOf Sequence Is IGenericCollection Then
        '@Ignore SetAssignmentWithIncompatibleObjectType
        Set Collection = Sequence
        
        If (Collection.Count = 0) Then _
             Call Me.Elements.Clear: Exit Function
                
        If (Collection.Count < MaxSize) Then _
            MaxSize = Collection.Count
    
    End If
    
    If TypeOf Sequence Is GenericSortedSet Then _
        If IsComparable(Me, Sequence) Then _
            Set Other = Sequence
                
    If (Other Is Nothing) Then

        With Sequence.Iterator
            
            If Not .HasNext(Element) Then _
                Exit Function
                
            ReDim SavedNodes.Elements(MaxSize - 1)
            
            If (This.Comparer Is Nothing) Then
            
                Do
                    Set Current = This.Root.Find(Element)
                    If Not (Current Is Nothing) Then
                        If Not ((Current.State And IS_MARKED) = IS_MARKED) Then
                            Set SavedNodes.Elements(SavedNodes.Size) = Current
                            SavedNodes.Size = SavedNodes.Size + 1
                            Current.State = (Current.State Or IS_MARKED) ' mark as already saved
                        End If
                    End If
                Loop While .HasNext(Element)
                
            Else
            
                Do
                    Set Current = This.Root.FindByComparer(Element, This.Comparer)
                    If Not (Current Is Nothing) Then
                        If Not ((Current.State And IS_MARKED) = IS_MARKED) Then
                            Set SavedNodes.Elements(SavedNodes.Size) = Current
                            SavedNodes.Size = SavedNodes.Size + 1
                            Current.State = (Current.State Or IS_MARKED) ' mark as already saved
                        End If
                    End If
                Loop While .HasNext(Element)
            
            End If
            
        End With
        
    Else 'fast Version
        
        Set OtherIterator = Other.Elements.Iterator
        otherEnded = Not OtherIterator.HasNext(Element)
        
        Call LoadStack(Stack, This.Root, False)
        thisEnded = Not InOrderTreeWalk(Current, Stack)
        
        Set Max = Me.GetMax
        ReDim SavedNodes.Elements(MaxSize - 1)
        
        If (This.Comparer Is Nothing) Then
            
            Do
                
                Set Value = Element 'Try cast
                If Value.IsGreaterThan(Max) Then _
                    Exit Do
                
                If Value.IsLessThan(Current.Element) Then
                    thisEnded = Not InOrderTreeWalk(Current, Stack)
                    
                ElseIf (Value.IsGreaterThan(Current.Element)) Then
                    otherEnded = Not OtherIterator.HasNext(Element)
                    
                Else
                    Set SavedNodes.Elements(SavedNodes.Size) = Current
                    SavedNodes.Size = SavedNodes.Size + 1
                    thisEnded = Not InOrderTreeWalk(Current, Stack)
                    otherEnded = Not OtherIterator.HasNext(Element)
                    
                End If
            
            Loop While (thisEnded = False And otherEnded = False)
        
        Else
            
            With This.Comparer
                Do
                     
                    If .AscendingOrdered(Max, Element) Then _
                        Exit Do
                    
                    If .AscendingOrdered(Element, Current.Element) Then
                        thisEnded = Not InOrderTreeWalk(Current, Stack)
                    
                    ElseIf .AscendingOrdered(Current.Element, Element) Then
                        otherEnded = Not OtherIterator.HasNext(Element)
                    
                    Else
                        Set SavedNodes.Elements(SavedNodes.Size) = Current
                        SavedNodes.Size = SavedNodes.Size + 1
                        thisEnded = Not InOrderTreeWalk(Current, Stack)
                        otherEnded = Not OtherIterator.HasNext(Element)
                        
                    End If
                
                Loop While (thisEnded = False And otherEnded = False)
            End With
            
        End If
    
    End If
    
    Set Current = Nothing
    
    If (SavedNodes.Size = This.Root.Count) Then 'current tree is already the intersection
        If (Other Is Nothing) Then 'roll back states
            Do While (SavedNodes.Size > 0) 'just in case
                SavedNodes.Size = SavedNodes.Size - 1
                Set Current = SavedNodes.Elements(SavedNodes.Size)
                Current.State = (Current.State Xor IS_MARKED)
            Loop
        End If
        Exit Function
    End If
    
    Call This.Root.Clear 'clear all left and right nodes
    Set This.Root = Nothing
    
    If (SavedNodes.Size = 0) Then _
        Exit Function
    
    SavedNodes.Size = SavedNodes.Size - 1
    Set This.Root = SavedNodes.Elements(SavedNodes.Size)
    This.Root.State = IS_BLACK
    This.Root.Count = 1
    
    If (This.Comparer Is Nothing) Then
    
        Do While (SavedNodes.Size > 0)
            
            SavedNodes.Size = SavedNodes.Size - 1
            Set SavedNode = SavedNodes.Elements(SavedNodes.Size)
            Set Current = This.Root
            Set Value = SavedNode.Element 'try cast
           
            With Value
                Do
                   
                    Call Current.AddCount(1)
                    
                    If Split4Node(Current) Then _
                        If Not (Parent Is Nothing) Then _
                            If (Parent.State = IS_RED) Then _
                                Call EnsureBalance(Current, Parent, GrandParent, GreatGrandParent)
                    
                    Set GreatGrandParent = GrandParent
                    Set GrandParent = Parent
                    Set Parent = Current
                    
                    If .IsLessThan(Current.Element) Then
                        IsLower = True
                        Set Current = Current.Left
                    Else
                        IsLower = False
                        Set Current = Current.Right
                    End If
            
                Loop While Not (Current Is Nothing)
            End With
            
            With SavedNode 'reset
                Set .Parent = Parent
                .State = IS_RED
                .Count = 1
            End With
            
            If IsLower Then
                Set Parent.Left = SavedNode
            Else
                Set Parent.Right = SavedNode
            End If
            
            If (Parent.State = IS_RED) Then _
                Call EnsureBalance(SavedNode, Parent, GrandParent, GreatGrandParent)
            
            '!clear stack
            Set GreatGrandParent = Nothing
            Set GrandParent = Nothing
            Set Parent = Nothing
                    
            This.Root.State = IS_BLACK
           
        Loop
    
    Else
        
        With This.Comparer
            Do While (SavedNodes.Size > 0)
                
                SavedNodes.Size = SavedNodes.Size - 1
                Set SavedNode = SavedNodes.Elements(SavedNodes.Size)
                Set Current = This.Root
                Set Element = SavedNode.Element
              
                Do
    
                    Call Current.AddCount(1)
                    
                    If Split4Node(Current) Then _
                        If Not (Parent Is Nothing) Then _
                            If (Parent.State = IS_RED) Then _
                                Call EnsureBalance(Current, Parent, GrandParent, GreatGrandParent)
                    
                    Set GreatGrandParent = GrandParent
                    Set GrandParent = Parent
                    Set Parent = Current
                    
                    If .AscendingOrdered(Element, Current.Element) Then
                        IsLower = True
                        Set Current = Current.Left
                    Else
                        IsLower = False
                        Set Current = Current.Right
                    End If
            
                Loop While Not (Current Is Nothing)

                With SavedNode 'reset
                    Set .Parent = Parent
                    .State = IS_RED
                    .Count = 1
                End With
                
                If IsLower Then
                    Set Parent.Left = SavedNode
                Else
                    Set Parent.Right = SavedNode
                End If
                
                If (Parent.State = IS_RED) Then _
                    Call EnsureBalance(SavedNode, Parent, GrandParent, GreatGrandParent)
                
                '!clear stack
                Set GreatGrandParent = Nothing
                Set GrandParent = Nothing
                Set Parent = Nothing
                        
                This.Root.State = IS_BLACK
                
            Loop
        End With
        
    End If
  
End Function

Public Function Overlaps(ByVal Sequence As IGenericSequence) As Boolean
    
    Dim Collection As IGenericCollection, Other As GenericSortedSet
    Dim Node As GenericTreeNode, Stack As InternalStack
    '@Ignore VariableNotAssigned
    Dim Candidate As Object
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Overlaps").IsNothing("Sequence")
    
    If (This.Root Is Nothing) Then _
        Exit Function
    
    If (Sequence Is Me) Then _
        Overlaps = True: Exit Function
    
    If TypeOf Sequence Is IGenericCollection Then
        '@Ignore SetAssignmentWithIncompatibleObjectType
        Set Collection = Sequence
        
        If (Collection.Count = 0) Then _
            Exit Function
                
    End If
    
    If TypeOf Sequence Is GenericSortedSet Then _
        If IsComparable(Me, Sequence) Then _
            Set Other = Sequence
                
    If (Other Is Nothing) Then
    
        With Sequence.Iterator
            If (This.Comparer Is Nothing) Then
            
                Do While .HasNext(Candidate)
                    If Not (This.Root.Find(Candidate) Is Nothing) Then _
                        Overlaps = True: Exit Function
                Loop
                
            Else
                
                Do While .HasNext(Candidate)
                    If Not (This.Root.FindByComparer(Candidate, This.Comparer) Is Nothing) Then _
                        Overlaps = True: Exit Function
                Loop
            
            End If
        End With
    
    Else
    
        Call LoadStack(Stack, This.Root, False)
      
        With Other
            Do While InOrderTreeWalk(Node, Stack)
                If .Contains(Node.Element) Then _
                    Overlaps = True: Exit Function
            Loop
        End With
    
    End If
    
End Function

Private Function Comparable(ByVal Element As IGenericValue) As IGenericValue: Set Comparable = Element: End Function

'TODO Check comparer Equality
Private Function IsComparable(ByVal A As GenericSortedSet, ByVal B As GenericSortedSet) As Boolean
    IsComparable = (A.Direction = B.Direction And A.Comparer Is B.Comparer)
End Function

Public Function TryRemove(ByVal Element As Object) As Boolean

    Call GenericError.Build(Me, "Public Function TryRemove").NotImplemented
    
    If (This.Root Is Nothing) Then _
        Exit Function
        
    If (This.Comparer Is Nothing) Then
        TryRemove = RemoveNode(This.Root.Find(Element))
    Else
        TryRemove = RemoveNode(This.Root.FindByComparer(Element, This.Comparer))
    End If
    
End Function

Public Function Contains(ByVal Element As Object) As Boolean
    
    If (This.Root Is Nothing) Then _
        Exit Function
        
    If (This.Comparer Is Nothing) Then
        Contains = ((This.Root.Find(Element) Is Nothing) = False)
    Else
        Contains = ((This.Root.FindByComparer(Element, This.Comparer) Is Nothing) = False)
    End If

End Function

Public Function ContainsAll(ByVal Sequence As IGenericSequence) As Boolean
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function ContainsAll").IsNothing("Sequence")
    
    If (This.Root Is Nothing) Then _
        Exit Function
    
    ContainsAll = This.Root.ContainsAll(Sequence, This.Comparer)

'@Ignore FunctionReturnValueAlwaysDiscarded
End Function

'@Ignore FunctionReturnValueAlwaysDiscarded
Public Function TryAdd(ByVal Element As Object) As Boolean
    
    Dim Node As GenericTreeNode, Current As GenericTreeNode
    Dim Parent  As GenericTreeNode, GrandParent As GenericTreeNode, GreatGrandParent As GenericTreeNode
    Dim IsLower As Boolean
    
    Dim Casted As IGenericValue
    
    If (This.Root Is Nothing) Then
        Set This.Root = New GenericTreeNode
        With This.Root
            Set .Element = Element
            .State = IS_BLACK
            .Count = 1
        End With
        TryAdd = True
        Exit Function
    End If
    
    Set Current = This.Root
    
    If (This.Comparer Is Nothing) Then
        Set Casted = Element 'Try cast
        With Casted
            Do
                
                If .Equals(Current.Element) Then
                    This.Root.State = IS_BLACK
                    Do While Parent Is Nothing = False
                        Call Parent.AddCount(-1)
                        Set Parent = Parent.Parent
                    Loop

                    Exit Function
                End If
                
                Call Current.AddCount(1)
                
                If Split4Node(Current) Then _
                    If Not (Parent Is Nothing) Then _
                        If (Parent.State = IS_RED) Then _
                            Call EnsureBalance(Current, Parent, GrandParent, GreatGrandParent)
    
                Set GreatGrandParent = GrandParent
                Set GrandParent = Parent
                Set Parent = Current
                
                If .IsLessThan(Current.Element) Then
                    Set Current = Current.Left
                    IsLower = True
                Else
                    Set Current = Current.Right
                    IsLower = False
                End If
    
            Loop While Not (Current Is Nothing)
        End With
        
    Else
    
        With This.Comparer
            Do
               
                If .Equality(Element, Current.Element) Then
                    This.Root.State = IS_BLACK
                    Do While Parent Is Nothing = False
                        Call Parent.AddCount(-1)
                        Set Parent = Parent.Parent
                    Loop
                    Exit Function
                End If
                
                Call Current.AddCount(1)
                
                If Split4Node(Current) Then _
                    If Not (Parent Is Nothing) Then _
                        If (Parent.State = IS_RED) Then _
                            Call EnsureBalance(Current, Parent, GrandParent, GreatGrandParent)
                
                Set GreatGrandParent = GrandParent
                Set GrandParent = Parent
                Set Parent = Current
                
                If .AscendingOrdered(Element, Current.Element) Then
                    Set Current = Current.Left
                    IsLower = True
                Else
                    Set Current = Current.Right
                    IsLower = False
                End If
        
            Loop While Not (Current Is Nothing)
            
        End With
    End If

    Set Node = New GenericTreeNode
    
    With Node
        Set .Element = Element
        Set .Parent = Parent
        .State = IS_RED
        .Count = 1
    End With
    
    If IsLower Then
        Set Parent.Left = Node
    Else
        Set Parent.Right = Node
    End If
    
    If (Parent.State = IS_RED) Then _
        Call EnsureBalance(Node, Parent, GrandParent, GreatGrandParent)
    
    This.Root.State = IS_BLACK
    TryAdd = True
     
End Function

Public Function Reverse() As IGenericIterator

    Dim Clone As Member

    If (This.Root Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequence_Iterator").IsNothing("(Root Is Nothing)")
    
    Call LoadStack(Clone.Iterator.Stack, This.Root, True)
    Clone.Iterator.HasLoaded = True
   
    Set Reverse = New GenericSortedSet
    Call System.Inject(Reverse, VarPtr(Clone), LenB(Clone))
    
End Function

Private Function FindMaxOf(ByVal Node As GenericTreeNode) As GenericTreeNode
    
    If Not (Node Is Nothing) Then _
        Do While (Node.Right Is Nothing) = False: Set Node = Node.Right: Loop
    
    Set FindMaxOf = Node

End Function

Private Function FindMinOf(ByVal Node As GenericTreeNode) As GenericTreeNode
    
    If Not (Node Is Nothing) Then _
        Do While ((Node.Left Is Nothing)) = False: Set Node = Node.Left: Loop
    
    Set FindMinOf = Node
    
End Function

Private Function InOrderNext(ByVal Node As GenericTreeNode) As GenericTreeNode
    
    Dim p As GenericTreeNode
    
    If (Node Is Nothing) Then _
        Exit Function
        
    If ((Node.Right Is Nothing) = False) Then
        Set p = FindMinOf(Node.Right)
    Else
        Set p = Node.Parent
        Do While ((p Is Nothing) = False)
            If ((p.Right Is Node) = False) Then _
                Exit Do
            
            Set Node = p
            Set p = p.Parent
        Loop
    End If
    
    Set InOrderNext = p
        
End Function

Private Function InOrderPrevious(ByVal Node As GenericTreeNode) As GenericTreeNode

    Dim p As GenericTreeNode
    
    If (Node Is Nothing) Then _
        Exit Function
        
    If ((Node.Left Is Nothing) = False) Then
        Set p = FindMaxOf(Node.Left)
    Else
        Set p = Node.Parent
        Do While ((p Is Nothing) = False)
            If ((p.Left Is Node) = False) Then _
                Exit Do
            
            Set Node = p
            Set p = p.Parent
        Loop
    End If
    
    Set InOrderPrevious = p

End Function
'TODO
Private Function RemoveNode(ByVal Node As GenericTreeNode) As Boolean

    Call GenericError.Build(Me, "Private Sub RemoveNode").NotImplemented
    
    If (Node Is Nothing) Then _
        Exit Function
    
    With Node
        Set .Parent = Nothing
        Set .Left = Nothing
        Set .Right = Nothing
    End With
    
    RemoveNode = True
    
End Function

Private Sub ReplaceNode( _
                            ByVal Match As GenericTreeNode, _
                            ByVal ParentOfMatch As GenericTreeNode, _
                            ByVal Succesor As GenericTreeNode, _
                            ByVal ParentOfSuccesor As GenericTreeNode)

    If (Succesor Is Match) Then
        If Not (Match.Right Is Nothing) Then _
            Call GenericError.Build(Me, "Private Sub ReplaceNode").InvalidOperation("Right child must be null!")

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

Private Sub EnsureBalance( _
                                ByVal Current As GenericTreeNode, _
                                ByRef Parent As GenericTreeNode, _
                                ByVal GrandParent As GenericTreeNode, _
                                ByVal GreatGrandParent As GenericTreeNode)

'    If (GrandParent Is Nothing) Then _
'        Call GenericError.Build(Me, "Private Sub InsertionBalance").NullPointer("Grand parent cannot be null here!")

    Dim CurrentIsOnRight As Boolean: CurrentIsOnRight = Parent.Right Is Current
    Dim NewChildOfGreatGrandParent As GenericTreeNode
    
    If ((GrandParent.Right Is Parent) = CurrentIsOnRight) Then
        If (CurrentIsOnRight) Then
            Set NewChildOfGreatGrandParent = GrandParent.RotateLeft
        Else
            Set NewChildOfGreatGrandParent = GrandParent.RotateRight
        End If
        
    Else
        If (CurrentIsOnRight) Then
            Set NewChildOfGreatGrandParent = GrandParent.RotateLeftRight
        Else
            Set NewChildOfGreatGrandParent = GrandParent.RotateRightLeft
        End If

        Set Parent = GreatGrandParent
    End If
    
    GrandParent.State = IS_RED
    NewChildOfGreatGrandParent.State = IS_BLACK
    
    If GreatGrandParent Is Nothing Then
        Set This.Root = NewChildOfGreatGrandParent
    Else
        If (GreatGrandParent.Left Is GrandParent) Then
            Set GreatGrandParent.Left = NewChildOfGreatGrandParent
        Else
            Set GreatGrandParent.Right = NewChildOfGreatGrandParent
        End If
    End If

End Sub

'Recursive
Private Sub AddToArray(ByVal Node As GenericTreeNode, ByVal Target As GenericArray, ByRef Index As Long)

    If Not (Node.Left Is Nothing) Then _
        Call AddToArray(Node.Left, Target, Index)
    
    Set Target.ElementAt(Index) = Node.Element: Index = Index + 1
    
    If Not (Node.Right Is Nothing) Then _
        Call AddToArray(Node.Right, Target, Index)
    
End Sub

Private Function CheckRedBlackBalance( _
                                            ByVal Node As GenericTreeNode, _
                                            Optional ByRef maxH As Long, _
                                            Optional ByRef minH As Long) As Boolean
    
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

    maxH = VBA.IIf(lmaxH < rmaxH, rmaxH, lmaxH) + 1
    minH = VBA.IIf(lminH > rminH, rminH, lminH) + 1
    
    If maxH <= 2 * minH Then _
        CheckRedBlackBalance = True
    
End Function

Private Function CheckHeightBalance(ByVal Node As GenericTreeNode) As Boolean

    Dim lh As Long
    Dim rh As Long
    
    If Node Is Nothing Then _
        CheckHeightBalance = True: Exit Function
        
    lh = HeightOf(Node.Left)
    rh = HeightOf(Node.Right)
    
    'High balanced
    If VBA.Abs(lh - rh) <= 1 Then _
        If CheckHeightBalance(Node.Left) Then _
            If CheckHeightBalance(Node.Left) Then _
                CheckHeightBalance = True
  
End Function

Private Function HeightOf(ByVal Node As GenericTreeNode) As Long
    
    Dim l As Long
    Dim r As Long
    
    If (Node Is Nothing) Then
        HeightOf = -1
        Exit Function
    End If
    
    l = HeightOf(Node.Left)
    r = HeightOf(Node.Right)
    HeightOf = VBA.IIf(l > r, l, r) + 1

End Function

Private Sub ReplaceChildOfNodeOrRoot( _
                                        ByVal Parent As GenericTreeNode, _
                                        ByVal Child As GenericTreeNode, _
                                        ByVal NewChild As GenericTreeNode)
    
    If Parent Is Nothing Then _
        Set This.Root = NewChild: Exit Sub
    
    If (Parent.Left Is Child) Then
        Set Parent.Left = NewChild
    Else
        Set Parent.Right = NewChild
    End If
        
End Sub

Private Function SizeOf(ByVal Node As GenericTreeNode) As Long:
    If Node Is Nothing = False Then _
        SizeOf = Node.Count
End Function

Private Function Split4Node(ByVal Node As GenericTreeNode) As Boolean
    
    If Node.HasLeaf Then _
        Exit Function
    
    If Node.Left.State = IS_BLACK Then _
        Exit Function
        
    If Node.Right.State = IS_BLACK Then _
        Exit Function
            
    With Node
        .State = IS_RED
        .Left.State = IS_BLACK
        .Right.State = IS_BLACK
    End With
    
    Split4Node = True
    
End Function

Private Function RotateLeft(ByVal x As GenericTreeNode) As GenericTreeNode

    Dim y  As GenericTreeNode
    Dim i As Long, j As Long
    
    With x
        Set y = .Right
        Set .Right = y.Left
        
        If Not .Right Is Nothing Then _
            Set .Right.Parent = x: j = .Right.Count
        
        Set y.Parent = .Parent
        Set y.Left = x: Set .Parent = y
        
        y.Count = .Count
        
        If Not .Left Is Nothing Then _
            i = .Left.Count

        .Count = 1 + i + j
    End With
    
    Set RotateLeft = y
    
End Function

Private Function RotateRight(ByVal x As GenericTreeNode) As GenericTreeNode

    Dim y  As GenericTreeNode
    Dim i As Long, j As Long
    
    With x
        Set y = .Left
        Set .Left = y.Right
        
        If Not .Left Is Nothing Then _
            Set .Left.Parent = x: i = .Left.Count
        
        Set y.Parent = .Parent
        Set y.Right = x: Set .Parent = y
        
        y.Count = .Count
        
        If Not .Right Is Nothing Then _
           j = .Right.Count
            
        .Count = 1 + i + j
    End With

    Set RotateRight = y
    
End Function

Private Function RotateLeftRight(ByVal Node As GenericTreeNode) As GenericTreeNode
    
    Set Node.Left = RotateLeft(Node.Left)
    Set RotateLeftRight = RotateRight(Node)

End Function

Private Function RotateRightLeft(ByVal Node As GenericTreeNode) As GenericTreeNode

    Set Node.Right = RotateRight(Node.Right)
    Set RotateRightLeft = RotateLeft(Node)

End Function

Private Sub UpdateSizeOf(ByVal Node As GenericTreeNode)

    Dim i As Long, j As Long
    Dim leftNode As GenericTreeNode, rightNode As GenericTreeNode
    
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
                                    ByRef Current As GenericTreeNode, _
                                    ByRef Stack As InternalStack) As Boolean
    
    Dim Node As GenericTreeNode

    If (Stack.Size = 0) Then _
        Set Current = Nothing: Exit Function
    
    With Stack
        Set Current = .Elements(.Size)
        Set .Elements(.Size) = Nothing
        .Size = .Size - 1
        
        If Not .Reverse Then
        
            Set Node = Current.Right
    
            Do While ((Node Is Nothing) = False)
                Set .Elements(InterlockedIncrement(.Size)) = Node
                Set Node = Node.Left
            Loop
            
        Else
        
            Set Node = Current.Left
    
            Do While ((Node Is Nothing) = False)
                Set .Elements(InterlockedIncrement(.Size)) = Node
                Set Node = Node.Right
            Loop
        End If
    End With
    
    InOrderTreeWalk = True

End Function

Private Function ReverseTreeWalk( _
                                    ByRef Current As GenericTreeNode, _
                                    ByRef Stack As InternalStack) As Boolean
    
    Dim Node As GenericTreeNode

    If (Stack.Size = 0) Then _
        Set Current = Nothing: Exit Function
    
    With Stack
        Set Current = .Elements(.Size)
        Set .Elements(.Size) = Nothing
        .Size = .Size - 1
        
        Set Node = Current.Left
    
        Do While ((Node Is Nothing) = False)
            Set .Elements(InterlockedIncrement(.Size)) = Node
            Set Node = Node.Right
        Loop

    End With
    
    ReverseTreeWalk = True

End Function

'You can calculate base-n logarithms for any number x by dividing the natural logarithm of x by the natural logarithm of n as follows: LOGn(x) = Log(x) / Log(n)
Private Function MaxHeightOf(ByVal Root As GenericTreeNode) As Long: MaxHeightOf = 2 * CLng((VBA.Log(SizeOf(Root) + 1) / LOG2)): End Function

'TODO Stack 0 based
Private Sub LoadStack(ByRef Stack As InternalStack, ByVal Node As GenericTreeNode, ByVal Reverse As Boolean)
    
    If (Node Is Nothing) Then _
        Exit Sub
        
    With Stack
        .Reverse = Reverse
        .Size = 0 '
        ReDim .Elements(1 To MaxHeightOf(Node))
        
        If Not .Reverse Then
            Do
                Set .Elements(InterlockedIncrement(.Size)) = Node
                Set Node = Node.Left
            Loop While Not (Node Is Nothing)
        Else
            Do
                Set .Elements(InterlockedIncrement(.Size)) = Node
                Set Node = Node.Right
            Loop While Not (Node Is Nothing)
        End If
    End With
    
End Sub

Private Function CloneTree(ByVal Node As GenericTreeNode) As GenericTreeNode
    
    Dim Clone As GenericTreeNode
    
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

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericSortedSet: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = GenericSortedSet: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(This): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(This): End Property
Private Property Get IGeneric_HashCode() As Long:  IGeneric_HashCode = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericSortedSet):  End Property

Private Sub Class_Terminate()
    If Not This.Root Is Nothing Then _
        Call This.Root.Clear
End Sub

Private Function IGeneric_Equals(ByVal Other As Object) As Boolean

    If (Other Is Nothing) Then _
        Exit Function
    
    If TypeOf Other Is GenericSortedSet Then _
        IGeneric_Equals = Me.Equals(Other)
        
End Function

Private Function IGeneric_ToString() As String
    With This
        IGeneric_ToString = _
                            "Class: " & VBA.TypeName$(Me) & vbNewLine & _
                            "Elements: " & Me.Elements.Count
    End With
End Function

'IGenericCollection
Private Property Get IGenericCollection_HasRandomAccess() As Boolean: IGenericCollection_HasRandomAccess = False: End Property
Private Property Get IGenericCollection_Count() As Long
    If Not (This.Root Is Nothing) Then _
        IGenericCollection_Count = This.Root.Count
End Property

Private Function IGenericCollection_Contains(ByVal Element As Object) As Boolean
    
    If (This.Root Is Nothing) Then _
        Exit Function
        
    If (This.Comparer Is Nothing) Then
        IGenericCollection_Contains = Not (This.Root.Find(Element) Is Nothing)
    Else
        IGenericCollection_Contains = Not (This.Root.FindByComparer(Element, This.Comparer) Is Nothing)
    End If

End Function

Private Function IGenericCollection_GetAt(ByVal Index As Long) As Object

    If (This.Root Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericCollection_GetAt").IndexOutofRange("(Index >= Count")
        
    If (Index >= This.Root.Count) Then _
        Call GenericError.Build(Me, "Private Function IGenericCollection_GetAt").IndexOutofRange("(Index >= Count")
    
    Set IGenericCollection_GetAt = This.Root.GetAt(Index).Element
    
End Function

Private Sub IGenericCollection_CopyTo(ByVal Target As GenericArray, ByVal Index As Long)
    
    If (Target Is Nothing) Then _
        Call GenericError.Build(Me, "Private Sub IGenericCollection_CopyTo").IsNothing("Target")
    
    If (Index < Target.LowerBound) Or (Index >= Target.Length) Then _
        Call GenericError.Build(Me, "Private Sub IGenericCollection_CopyTo").IndexOutofRange("(Index < Target.LowerBound) Or (Index >= Target.Length)")
    
    Dim Stack As InternalStack
    Dim Current As GenericTreeNode
    
    Call LoadStack(Stack, This.Root, False)
    
    Index = Index - 1
    
    With Target
        Do While InOrderTreeWalk(Current, Stack): Set .ElementAt(InterlockedIncrement(Index)) = Current.Element: Loop
    End With

End Sub

Private Function IGenericCollection_ToArray() As GenericArray

    Set IGenericCollection_ToArray = GenericArray 'avoid null
    
    If (This.Root Is Nothing) Then _
        Exit Function
        
    Set IGenericCollection_ToArray = GenericArray.Build(This.Root.Count)
    Call Me.Elements.CopyTo(IGenericCollection_ToArray, IGenericCollection_ToArray.LowerBound)
    
End Function

Private Sub IGenericCollection_Clear()
    
    If Not (This.Root Is Nothing) Then _
        Call This.Root.Clear
        
End Sub

Private Function IGenericCollection_Copy() As IGenericCollection
    
    Dim Clone As Member
    
    Set Clone.Comparer = This.Comparer
    Set Clone.Root = CloneTree(This.Root)
    
    Set IGenericCollection_Copy = New GenericSortedSet
    Call System.Inject(IGenericCollection_Copy, VarPtr(Clone), LenB(Clone))
    
End Function

Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = IGenericSequence_Iterator: End Function

'IGenericSequence
Private Function IGenericSequence_Iterator() As IGenericIterator

    Dim Clone As Member

'    If (This.Root Is Nothing) Then _
'        Call GenericError.Build(Me, "Private Function IGenericSequence_Iterator").IsNothing("(Root Is Nothing)")
'
    Call LoadStack(Clone.Iterator.Stack, This.Root, False)
    Set Clone.Iterator.SortedSet = Me
    
    Clone.Iterator.HasLoaded = True
   
    Set IGenericSequence_Iterator = New GenericSortedSet
    Call System.Inject(IGenericSequence_Iterator, VarPtr(Clone), LenB(Clone))
    
End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As Object: Set IGenericIterator_Current = This.Iterator.Current: End Property

Private Function IGenericIterator_HasNext(Optional ByRef Out As Object) As Boolean
    
    Dim Node As GenericTreeNode

    If InOrderTreeWalk(Node, This.Iterator.Stack) Then
    
        Set This.Iterator.Current = Node.Element
        Set Out = This.Iterator.Current
        IGenericIterator_HasNext = True
        Exit Function
        
    End If
    
    If This.Iterator.HasLoaded = False Then _
        Call GenericError.Build(Me, "Private Function IGenericIterator_HasNext").InvalidIterator

    Set This.Iterator.Current = Nothing
    Erase This.Iterator.Stack.Elements
 
End Function

'IGenericPredicator
Private Function IGenericPredicator_IsTrue(ByVal A As Object, ByVal B As Object) As Boolean

    Dim First As GenericSortedSet, Second As GenericSortedSet
    Dim FirstIterator As IGenericIterator, SecondIterator As IGenericIterator
    
    Dim Comparer As IGenericComparer
    
    '@Ignore VariableNotAssigned
    '@Ignore VariableNotAssigned
    Dim x As IGenericValue, y As IGenericValue
    '@Ignore VariableNotAssigned
    '@Ignore VariableNotAssigned
    Dim ai As Object, bi As Object
    
    Dim FirstEnded As Boolean, SecondEnded As Boolean
    
    If A Is B Then _
        IGenericPredicator_IsTrue = True: Exit Function
        
    If A Is Nothing Then _
        Exit Function
    
    If B Is Nothing Then _
        Exit Function
   
    Set First = A: Set Second = B 'Try cast

    If Not (First.Elements.Count = Second.Elements.Count) Then _
        Exit Function
    
    If Not IsComparable(First, Second) Then _
        Exit Function
        
    Set FirstIterator = First.Elements.Iterator
    Set SecondIterator = Second.Elements.Iterator
    
    Set Comparer = First.Comparer
    If (Comparer Is Nothing) Then
         
         FirstEnded = Not FirstIterator.HasNext(x)
         SecondEnded = Not SecondIterator.HasNext(y)
         
        Do While (Not FirstEnded And Not SecondEnded)
            
            If Not x.Equals(y) Then _
                Exit Function
        
            FirstEnded = Not FirstIterator.HasNext(x)
            SecondEnded = Not SecondIterator.HasNext(y)
            
        Loop
    
    Else
            
        FirstEnded = Not FirstIterator.HasNext(ai)
        SecondEnded = Not SecondIterator.HasNext(bi)
         
        With Comparer
            Do While (Not FirstEnded And Not SecondEnded)
            
                If Not .Equality(ai, bi) Then _
                    Exit Function
                    
                FirstEnded = Not FirstIterator.HasNext(ai)
                SecondEnded = Not SecondIterator.HasNext(bi)
            Loop
        End With
        
    End If
   
    IGenericPredicator_IsTrue = (FirstEnded And SecondEnded) 'Both must be true

End Function


