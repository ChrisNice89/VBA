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
Option Compare Database

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
    Current As IGeneric
    Type As DictionaryData
    Delegate As Skynet
    Stack As InternalStack
    MaxStack As Long
    HasLoaded As Boolean
End Type

Private Type Member
    Root As GenericTreeNode
    Count As Long
    Iterator As Iterator
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get MaxHeight() As Long: MaxHeight = 2 * Skynet.Logn(myClass.Count + 1, Base:=2): End Property
Public Property Get Count() As Long: Count = myClass.Count: End Property

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

Public Function GetMax() As IGeneric
    
    Dim Node As GenericTreeNode: Set Node = myClass.Root
    
    If Not (Node Is Nothing) Then
        Do While Not (Node.Right Is Nothing)
            Set Node = Node.Right
        Loop
    End If
    
    Set GetMax = Node.AsPair

End Function

Public Function GetMin() As IGeneric
    
    Dim Node As GenericTreeNode: Set Node = myClass.Root
    
    If Not (Node Is Nothing) Then
        Do While Not (Node.Left Is Nothing)
            Set Node = Node.Left
        Loop
    End If
    
    Set GetMin = Node.AsPair

End Function

Public Sub AddAll(ByVal collection As IGenericIterator)
    
'    Dim p As GenericPair
'    Dim i As IGeneric
'
'    With Collection
'        Do While Collection.MoveNext
'            Set i = .Current
'            If TypeOf i Is GenericPair Then
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
    
    If (Me Is GenericTree) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption
    
    Set Build = New GenericTree 'Skynet.Create(GenericTree)

End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Virtual.Size = LenB(myClass)
End Sub

Public Function Iterator(ByVal T As DictionaryData, Optional ByVal Reverse As Boolean) As IGenericIterator
    
    Dim C As Member

    With C
        Set .Root = myClass.Root
        .Count = myClass.Count
        Call LoadIterator(.Iterator, T, Reverse)
    End With
    
    Set Iterator = Skynet.CreateInstance(New GenericTree, VarPtr(C), LenB(C))
    
End Function

Public Sub Clear()
    With myClass
        Set .Root = Nothing
        .Count = 0
    End With
End Sub

Public Function Remove(ByVal Item As IGeneric) As Boolean


End Function

Public Function ContainsKey(ByVal Key As IGeneric) As Boolean: ContainsKey = Not FindNode(Key) Is Nothing: End Function
Public Function ContainsValue(ByVal Value As IGeneric) As Boolean
    
    Dim Stack As InternalStack
    Dim Current As GenericTreeNode

    Call LoadStack(Stack, False)
    
    If Not (Value Is Nothing) Then
        Do While InOrderTreeWalk(Current, Stack)
            If (Current.Value.Equals(Value)) Then _
                ContainsValue = True: Exit Function
        Loop
    Else
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
    
    Dim Current As GenericTreeNode: Set Current = myClass.Root
    Dim Order As ComparerEnum
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function FindNode").KeyIsNothing
    
    Do While (Current Is Nothing) = False
        Order = Current.Key.CompareTo(Key)
        
        If (Order = IsEqual) Then _
            Set FindNode = Current: Exit Function
        
        If Order = IsLower Then
            Set Current = Current.Right
        Else
            Set Current = Current.Left
        End If
        
    Loop
    
End Function

Public Function IndexOf(ByVal Key As IGeneric) As Long
    
    Dim Node As GenericTreeNode: Set Node = myClass.Root
    
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
    Dim Order As Long
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Add").KeyIsNothing
    
    If (myClass.Root Is Nothing) Then _
        Set myClass.Root = GenericTreeNode(Key, Value, False): _
        Call InterlockedIncrement(myClass.Count): _
        Add = True: _
        Exit Function
    
    Set Current = myClass.Root
    Do
        Order = Current.Key.CompareTo(Key)
        
        If (Order = 0) Then _
            myClass.Root.IsRed = False: _
            Add = False: _
            Exit Function

        If Split4Node(Current) Then _
            If (Parent Is Nothing) = False Then _
                If Parent.IsRed Then _
                    Call InsertionBalance(Current, Parent, GrandParent, GreatGrandParent)
            
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

'    Set Node = GenericTreeNode(Key, Value)
    Set Node = New GenericTreeNode
    
    With Node
        Set .Key = Key: Set .Value = Value
        .IsRed = True
    End With
    
    If (Order < 0) Then
        Set Parent.Right = Node
    Else
        Set Parent.Left = Node
    End If
    
    If (Parent.IsRed) Then _
        Call InsertionBalance(Node, Parent, GrandParent, GreatGrandParent)
    
    myClass.Root.IsRed = False
    myClass.Count = myClass.Count + 1
    
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
        Set myClass.Root = NewChildOfGreatGrandParent
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
        Set myClass.Root = NewChild: Exit Sub
    
    If (Parent.Left Is Child) Then
        Set Parent.Left = NewChild
    Else
        Set Parent.Right = NewChild
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


Private Function RotateLeft(ByVal Node As GenericTreeNode) As GenericTreeNode

    Dim x  As GenericTreeNode: Set x = Node.Right
   
    Set Node.Right = x.Left: Set x.Left = Node
    
    Set RotateLeft = x
    
End Function

Private Function RotateLeftRight(ByVal Node As GenericTreeNode) As GenericTreeNode

    Dim Child  As GenericTreeNode: Set Child = Node.Left
    Dim GrandChild  As GenericTreeNode: Set GrandChild = Child.Right
    
    With GrandChild
        Set Node.Left = .Right: Set .Right = Node
        Set Child.Right = .Left: Set .Left = Child
    End With
   
    Set RotateLeftRight = GrandChild
    
End Function

Private Function RotateRight(ByVal Node As GenericTreeNode) As GenericTreeNode

    Dim x  As GenericTreeNode: Set x = Node.Left
    
    Set Node.Left = x.Right: Set x.Right = Node
    
    Set RotateRight = x
    
End Function

Private Function RotateRightLeft(ByVal Node As GenericTreeNode) As GenericTreeNode

    Dim Child  As GenericTreeNode: Set Child = Node.Right
    Dim GrandChild  As GenericTreeNode: Set GrandChild = Child.Left
    
    With GrandChild
        Set Node.Right = .Left: Set .Left = Node
        Set Child.Left = .Right: Set .Right = Child
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
    
    Dim Node As GenericTreeNode: Set Node = myClass.Root
    
    With Stack
        .Reverse = Reverse
        .Size = 0
        ReDim .Items(1 To MaxHeight)
        
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

Private Sub LoadIterator(ByRef E As Iterator, ByVal T As DictionaryData, ByVal Reverse As Boolean)
    
    With E
        .Type = T
        Call LoadStack(.Stack, Reverse)
        Set .Current = Nothing
        .HasLoaded = True
    End With
    
End Sub

Private Function CopyOf(ByVal T As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long, ByVal Caller As String) As GenericArray
    
    If (GArray Is Nothing) Then _
        Call Skynet.Throw(Me, Caller).NullPointer("GArray Is Nothing")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call Skynet.Throw(Me, Caller).IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < myClass.Count) Then _
        Call Skynet.Throw(Me, Caller).InvalidInput("(GArray.Length - Index < myClass.Count)")
    
    Dim Stack As InternalStack
    Dim Node As GenericTreeNode
    Dim Current As IGeneric
    Call LoadStack(Stack, False)

    Select Case True
        Case T = keys_
            Do While InOrderTreeWalk(Node, Stack)
                Set GArray(ArrayIndex) = Node.Key
                ArrayIndex = ArrayIndex + 1
            Loop
        Case T = Values_
            Do While InOrderTreeWalk(Node, Stack)
                Set GArray(ArrayIndex) = Node.Value
                ArrayIndex = ArrayIndex + 1
            Loop
        Case Else
            Do While InOrderTreeWalk(Node, Stack)
                Set GArray(ArrayIndex) = Node.AsPair
                ArrayIndex = ArrayIndex + 1
            Loop
    End Select

    Set CopyOf = GArray
    
End Function

'IGeneric
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = IGeneric_HashValue = ObjPtr(Me): End Property
Private Function IGeneric_Clone() As IGeneric
    
    Dim C As Member
    
    With C
        Set .Root = myClass.Root
        .Count = myClass.Count
    End With
    
    Set IGeneric_Clone = Skynet.CreateInstance(New GenericTree, VarPtr(C), LenB(C))
    
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
Private Property Set IGenericDictionary_Item(ByVal Key As IGeneric, ByVal Value As IGeneric): Set Me.Item(Key) = Value: End Property
Private Property Get IGenericDictionary_Item(ByVal Key As IGeneric) As IGeneric: Set IGenericDictionary_Item = Me.Item(Key): End Property
Private Property Get IGenericDictionary_Count() As Long: IGenericDictionary_Count = myClass.Count: End Property
Private Function IGenericDictionary_ContainsKey(ByVal Key As IGeneric) As Boolean: IGenericDictionary_ContainsKey = Me.ContainsKey(Key:=Key): End Function
Private Function IGenericDictionary_ContainsValue(ByVal Value As IGeneric) As Boolean: IGenericDictionary_ContainsValue = Me.ContainsValue(Value:=Value): End Function
Private Function IGenericDictionary_TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: IGenericDictionary_TryGetValue = Me.TryGetValue(Key, Value): End Function
Private Sub IGenericDictionary_Remove(ByVal Key As IGeneric): Call Me.Remove(Item:=Key): End Sub
Private Sub IGenericDictionary_Add(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value): End Sub
Private Sub IGenericDictionary_CopyOf(ByVal T As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(T:=T, GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="Private Sub IGenericDictionary_CopyOf"): End Sub
Private Function IGenericDictionary_Iterator(ByVal T As DictionaryData) As IGenericIterator: Set IGenericDictionary_Iterator = Me.Iterator(T:=T): End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = False: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = False: End Property
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.ContainsKey(Key:=Item): End Function
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = CopyOf(T:=Pairs_, GArray:=GenericArray.Build(myClass.Count), ArrayIndex:=GenericArray.LowerBound, Caller:="Private Function IGenericCollection_ToArray() As GenericArray"): End Function
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(T:=Pairs_, GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="Private Sub IGenericCollection_CopyTo"): End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Iterator(T:=Pairs_, Reverse:=False): End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myClass.Iterator.Current: End Property

Private Function IGenericIterator_MoveNext() As Boolean
    
    Dim Node As GenericTreeNode
    
    With myClass.Iterator
        If InOrderTreeWalk(Node, .Stack) Then
            
            IGenericIterator_MoveNext = True
            
            If .Type = keys_ Then _
                Set .Current = Node.Key: Exit Function
            
            If .Type = Values_ Then _
                Set .Current = Node.Value: Exit Function
            
            Set .Current = Node.AsPair: Exit Function
           
        End If
        
        If .HasLoaded = False Then _
            Call Skynet.Throw(Me, "Private Function IGenericIterator_MoveNext").InvalidIterator
          
        Set .Current = Nothing
      
    End With
 
End Function

Private Sub IGenericIterator_Reset()
    Call LoadIterator(myClass.Iterator, myClass.Iterator.Type, myClass.Iterator.Stack.Reverse)
End Sub

'IPrototype
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericTree: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GenericTree: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.Size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property



