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

Implements IGeneric
Implements IGenericDictionary
Implements IGenericCollection
Implements IGenericEnumerator
Implements IPrototype

Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Var() As Any) As Long

Private Const MAX_CALLSTACK_SIZE As Long = 1000

Private Type InternalStack
    Size As Long
    Items() As GenericTreeNode
    Reverse As Boolean
End Type

Private Type Enumerator
    Current As IGeneric
    Stack As InternalStack
    Parent As GenericTree
    MaxStack As Long
    HasLoaded As Boolean
End Type

Private Type Member
    Root As GenericTreeNode
    Count As Long
    Enumerator As Enumerator
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get MaxHeight() As Long: MaxHeight = 2 * Skynet.Logn(myClass.Count + 1, 2): End Property
Public Property Get Count() As Long: Count = myClass.Count: End Property

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
    
    
    Set GetMax = Node.CopyToPair

End Function

Public Function GetMin() As IGeneric
    
    Dim Node As GenericTreeNode: Set Node = myClass.Root
    
    If Not (Node Is Nothing) Then
        Do While Not (Node.Left Is Nothing)
            Set Node = Node.Left
        Loop
    End If
    
    Set GetMin = Node.CopyToPair

End Function

Public Sub AddAll(ByVal Collection As IGenericEnumerator)
    
'    Dim p As GenericPairXY
'    Dim i As IGeneric
'
'    With Collection
'        Do While Collection.MoveNext
'            Set i = .Current
'            If TypeOf i Is GenericPairXY Then
'                Set p = .Current
'                If Not Me.Contains(p.Key) Then _
'                    Call Me.Add(p.Key, p.Value)
'            Else
'                 If Not Me.Contains(.Current) Then _
'                    Call Me.Add(p.Key, p.Value)
'
'            End If
'
'        Loop
'    End With
'
End Sub

Public Function GetValues(Optional ByVal Reverse As Boolean) As GenericArray
   
    ReDim Values(GenericArray.LowerBound To myClass.Count) As IGeneric
    Dim Index As Long: Index = GenericArray.LowerBound

    Dim Stack As InternalStack
    Dim Current As GenericTreeNode

    Call LoadStack(Stack, Reverse)

    Do While InOrderTreeWalk(Current, Stack)
        Set Values(Index) = Current.Key
        Index = Index + 1
    Loop
    
    Set GetValues = GenericArray.BuildFrom(Values)
    
End Function

Public Function GetKeys(Optional ByVal Reverse As Boolean) As GenericArray
   
    ReDim Keys(GenericArray.LowerBound To myClass.Count) As IGeneric
    Dim Index As Long: Index = GenericArray.LowerBound

    Dim Stack As InternalStack
    Dim Current As GenericTreeNode

    Call LoadStack(Stack, Reverse)

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

Public Function Build() As GenericTree

   'Set myClass.Root = GenericTreeNode.Build(Nothing, Nothing)
    Set Build = New GenericTree 'Skynet.Create(GenericTree)

End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Virtual.Size = LenB(myClass)
End Sub

Public Function Enumerator(Optional ByVal Reverse As Boolean) As IGenericEnumerator
    
    Dim C As Member
     
    Call LoadEnumerator(C.Enumerator, Reverse)
    Set C.Enumerator.Parent = Me
    
    Set Enumerator = Skynet.CreateInstance(New GenericTree, VarPtr(C), Virtual.Size)
    
End Function

Public Sub Clear()
    Set myClass.Root = Nothing
    myClass.Count = 0
End Sub

Public Function Contains(ByVal Key As IGeneric) As Boolean: Contains = Not FindNode(Key) Is Nothing: End Function
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
    
        Set Current = IIf(Order < 0, Current.Right, Current.Left)
    Loop
    
End Function

Public Function IndexOf(ByVal Key As IGeneric) As Long
    
    Dim Node As GenericTreeNode: Set Node = myClass.Root
    
    Dim Count As Long: Count = 1
    Dim Order As Long
    
    Do While Not (Node Is Nothing)
        Order = Node.Key.CompareTo(Key)
        If (Order = 0) Then
           IndexOf = Count: Exit Function
        End If
        
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
    Dim Order As Long
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Add").KeyIsNothing
    
    If (myClass.Root Is Nothing) Then
        Set myClass.Root = GenericTreeNode(Key, Value, False)
        Call InterlockedIncrement(myClass.Count)
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
             'Call Split4Node(Current)
            With Current
                .IsRed = True
                .Left.IsRed = False
                .Right.IsRed = False
            End With
            
            If Not Parent Is Nothing Then _
                If Parent.IsRed Then _
                    Call InsertionBalance(Current, Parent, GrandParent, GreatGrandParent)

        End If
        
        Set GreatGrandParent = GrandParent
        Set GrandParent = Parent
        Set Parent = Current
        Set Current = IIf(Order > 0, Current.Left, Current.Right)
     
   Loop
   
'    If (Parent Is Nothing) Then _
'        Call Skynet.Throw(Me, "Private Sub InsertionBalance").NullPointer("Parent cannot be null here!")

    Set Node = GenericTreeNode(Key, Value)
    
    If (Order < 0) Then
        Set Parent.Right = Node
    Else
        Set Parent.Left = Node
    End If
    
    If (Parent.IsRed) Then _
        Call InsertionBalance(Node, Parent, GrandParent, GreatGrandParent)
    
    myClass.Root.IsRed = False
    InterlockedIncrement myClass.Count
    
    Add = True
            
End Function

Private Sub InsertionBalance( _
                                ByVal Current As GenericTreeNode, _
                                ByRef Parent As GenericTreeNode, _
                                ByVal GrandParent As GenericTreeNode, _
                                ByVal GreatGrandParent As GenericTreeNode)

            
'    If (GrandParent Is Nothing) Then _
'        Call Skynet.Throw(Me, "Private Sub InsertionBalance").NullPointer("Grand parent cannot be null here!")

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
    Set Node.Left = x.Right
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

Private Sub LoadStack(ByRef Stack As InternalStack, ByVal Reverse As Boolean)
    
    Dim Node As GenericTreeNode
    Set Node = myClass.Root
    
    With Stack
        .Reverse = Reverse
        .Size = 0
        ReDim .Items(1 To Me.MaxHeight)
        
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

Private Sub LoadEnumerator(ByRef e As Enumerator, ByVal Reverse As Boolean)
    
    With e
        Call LoadStack(.Stack, Reverse)
        Set .Current = Nothing
        .HasLoaded = True
    End With
    
End Sub

'IGeneric
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = IGeneric_HashValue = ObjPtr(Me): End Property
Private Function IGeneric_Clone() As IGeneric
    
    Dim C As Member: C = myClass
    Call Skynet.Throw(Me, "Private Function IGeneric_Equals").NotImplemented
    Set IGeneric_Clone = Skynet.CreateInstance(New GenericTree, VarPtr(C), Virtual.Size)
  
End Function

Private Sub IGeneric_Dispose()
    myClass = EmptyStructure
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
    With myClass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Elements: " & .Count
    End With
End Function

Private Sub Class_Terminate()
    Call Skynet.Dispose(Me)
End Sub

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
    
    If (GArray Is Nothing) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericDictionary_CopyKeysTo").NullPointer("GArray Is Nothing")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericDictionary_CopyKeysTo").IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < Me.Count) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericDictionary_CopyKeysTo").InvalidInput("(GArray.Length - Index < Me.Count)")
    
    Dim Stack As InternalStack
    Dim Current As GenericTreeNode
    
    Call LoadStack(Stack, False)

    Do While InOrderTreeWalk(Current, Stack)
        Set GArray(ArrayIndex) = Current.Key
        ArrayIndex = ArrayIndex + 1
    Loop
  
End Sub

Private Sub IGenericDictionary_CopyValuesTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long)
    
    If (GArray Is Nothing) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericDictionary_CopyValuesTo").NullPointer("GArray Is Nothing")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericDictionary_CopyValuesTo").IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < Me.Count) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericDictionary_CopyValuesTo").InvalidInput("(GArray.Length - Index < Me.Count)")
    
    Dim Stack As InternalStack
    Dim Current As GenericTreeNode
    
    Call LoadStack(Stack, False)

    Do While InOrderTreeWalk(Current, Stack)
        Set GArray(ArrayIndex) = Current.Value
        ArrayIndex = ArrayIndex + 1
    Loop
    
End Sub

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = False: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = False: End Property
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.Contains(Key:=Item): End Function
Private Function IGenericCollection_ToArray() As GenericArray

    Dim GArray As GenericArray
    Set GArray = GenericArray.Build(myClass.Count)
   
    Call IGenericCollection_CopyTo(GArray:=GArray, ArrayIndex:=GArray.LowerBound)
    
    Set IGenericCollection_ToArray = GArray

End Function

Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long)

    If (GArray Is Nothing) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericCollection_CopyTo").NullPointer("GArray Is Nothing")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericCollection_CopyTo").IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < Me.Count) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericCollection_CopyTo").InvalidInput("(GArray.Length - Index < Me.Count)")
    
    Dim Current As GenericTreeNode
    Dim Stack As InternalStack
    
    Call LoadStack(Stack, False)
    
    Do While InOrderTreeWalk(Current, Stack)
        Set GArray(ArrayIndex) = Current.CopyToPair
        ArrayIndex = ArrayIndex + 1
    Loop
    
End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_NewEnumerator() As IGenericEnumerator: Set IGenericCollection_NewEnumerator = Me.Enumerator: End Function

'IGenericEnumerator
Private Property Get IGenericEnumerator_Current() As IGeneric: Set IGenericEnumerator_Current = myClass.Enumerator.Current: End Property

Private Function IGenericEnumerator_MoveNext() As Boolean
    
    Dim Node As GenericTreeNode
    
    If InOrderTreeWalk(Node, myClass.Enumerator.Stack) Then _
        Set myClass.Enumerator.Current = Node.CopyToPair: IGenericEnumerator_MoveNext = True
    'case Enumerator is  GenericTree
    If myClass.Enumerator.HasLoaded = False Then _
        Call LoadEnumerator(myClass.Enumerator, myClass.Enumerator.Stack.Reverse): _
        If InOrderTreeWalk(Node, myClass.Enumerator.Stack) Then _
            Set myClass.Enumerator.Current = Node.CopyToPair: IGenericEnumerator_MoveNext = True
 
End Function

Private Sub IGenericEnumerator_Reset()
    Call LoadEnumerator(myClass.Enumerator, myClass.Enumerator.Stack.Reverse)
End Sub

'IPrototype
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericTree: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GenericTree: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.Size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property



