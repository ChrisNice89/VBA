VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericTreeNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Collections"
Option Explicit

'Performance
'@Ignore EncapsulatePublicField
Public State As Byte
'@Ignore EncapsulatePublicField
Public Count As Long
'@Ignore EncapsulatePublicField
Public Element As Object
'@Ignore EncapsulatePublicField
Public Parent As GenericTreeNode
'@Ignore EncapsulatePublicField
Public Left As GenericTreeNode
'@Ignore EncapsulatePublicField
Public Right As GenericTreeNode

Public Property Get Instance() As IGeneric: Set Instance = Me: End Property

Public Property Get IsLeaf() As Boolean
    If Left Is Nothing Then _
        If Right Is Nothing Then _
            IsLeaf = True
End Property

Public Property Get HasLeaf() As Boolean
    
    If Left Is Nothing Then _
        HasLeaf = True: Exit Property
    
    If Right Is Nothing Then _
        HasLeaf = True: Exit Property
        
End Property

'Private Sub Class_Terminate()
'    Debug.Print "Node Delete"
'End Sub

Public Sub AddCount(ByVal i As Long): Count = Count + i: End Sub

Public Function Copy() As GenericTreeNode
    
    Dim newNode As GenericTreeNode
    Set newNode = New GenericTreeNode
    
    With newNode
        Set .Element = Element
        .State = State
        .Count = Count
    End With
    
    Set Copy = newNode
    
End Function

Public Function After() As GenericTreeNode
    
    Dim p As GenericTreeNode
    Dim n As GenericTreeNode
    
    If Right Is Nothing = False Then
        Set p = Right
        Do While (p.Left Is Nothing) = False
            Set p = p.Left
        Loop
    Else
        Set p = Parent
        Set n = Me
        Do While (p Is Nothing) = False
            If (p.Right Is n) = False Then _
                Exit Do
            
            Set n = p
            Set p = p.Parent
        Loop
    End If
    
    Set After = p
        
End Function

Public Function Before() As GenericTreeNode
    
    Dim p As GenericTreeNode
    Dim n As GenericTreeNode
    
    If Left Is Nothing = False Then
        Set p = Left
        Do While (p.Right Is Nothing) = False
            Set p = p.Right
        Loop
    Else
        Set p = Parent
        Set n = Me
        Do While (p Is Nothing) = False
            If (p.Left Is n) = False Then _
                Exit Do
            
            Set n = p
            Set p = p.Parent
        Loop
    End If
    
    Set Before = p
        
End Function

Public Function RotateLeft() As GenericTreeNode

    Dim y  As GenericTreeNode
    Dim i As Long, j As Long
    
    Set y = Right
    Set Right = y.Left
    
    If Not Right Is Nothing Then _
        Set Right.Parent = Me: j = Right.Count
    
    Set y.Parent = Parent
    Set y.Left = Me
    Set Parent = y
    
    y.Count = Count

    If Not Left Is Nothing Then _
        i = Left.Count

    Count = 1 + i + j

    Set RotateLeft = y
    
End Function

Public Function RotateRight() As GenericTreeNode

    Dim y  As GenericTreeNode
    Dim i As Long, j As Long
  
    Set y = Left
    Set Left = y.Right
    
    If Not Left Is Nothing Then _
        Set Left.Parent = Me: i = Left.Count
    
    Set y.Parent = Parent
    Set y.Right = Me
    Set Parent = y
    
    y.Count = Count

    If Not Right Is Nothing Then _
       j = Right.Count

    Count = 1 + i + j
        
    Set RotateRight = y
    
End Function

Public Function RotateLeftRight() As GenericTreeNode: Set Left = Left.RotateLeft: Set RotateLeftRight = Me.RotateRight: End Function
Public Function RotateRightLeft() As GenericTreeNode: Set Right = Right.RotateRight: Set RotateRightLeft = Me.RotateLeft: End Function

Public Function FindByComparer(ByVal Element As Object, ByVal Comparer As IGenericComparer) As GenericTreeNode
    
    Set FindByComparer = Me
    With Comparer
        Do
        
            If .Equality(Element, FindByComparer.Element) Then _
                Exit Function
                
            If .AscendingOrdered(Element, FindByComparer.Element) Then
                Set FindByComparer = FindByComparer.Left
            Else
                Set FindByComparer = FindByComparer.Right
            End If
        
        Loop While Not (FindByComparer Is Nothing)
    End With
    
End Function

Public Function Find(ByVal Element As IGenericValue) As GenericTreeNode
       
    Set Find = Me
    With Element
        Do
        
            If .Equals(Find.Element) Then _
                Exit Function
            
            If .IsLessThan(Find.Element) Then
                Set Find = Find.Left
            Else
                Set Find = Find.Right
            End If

        Loop While Not (Find Is Nothing)
    End With
    
End Function

Public Function GetAt(ByVal Rank As Long) As GenericTreeNode
    
    Dim leftSize As Long
    
    Dim Node As GenericTreeNode
    Set Node = Me
    
    Do
        
        If ((Node.Left Is Nothing) = False) Then
            leftSize = Node.Left.Count
        Else
            leftSize = 0
        End If
        
        If (Rank = leftSize) Then _
            Exit Do
      
        If (leftSize > Rank) Then
            Set Node = Node.Left
        Else
            Set Node = Node.Right
            Rank = Rank - leftSize - 1
        End If
        
    Loop While Not (Node Is Nothing)
    
    Set GetAt = Node
    
End Function

Public Function GetMax() As GenericTreeNode
    
    Dim Node As GenericTreeNode, Candidate As GenericTreeNode
    
    Set Node = Me
    Set Candidate = Node.Right
    
    Do While Not (Candidate Is Nothing)
        Set Node = Candidate
        Set Candidate = Node.Right
    Loop
    
    Set GetMax = Node

End Function

Public Function GetMin() As GenericTreeNode
    
    Dim Node As GenericTreeNode, Candidate As GenericTreeNode
    
    Set Node = Me
    Set Candidate = Node.Left
    
    Do While Not (Candidate Is Nothing)
        Set Node = Candidate
        Set Candidate = Node.Left
    Loop
    
    Set GetMin = Node

End Function

Public Function GetAfter(ByVal Element As Object, Optional ByVal Comparer As IGenericComparer = Nothing) As GenericTreeNode
    
    Dim Current As GenericTreeNode
    Dim BestSoFar As GenericTreeNode
    Dim Pivot As IGenericValue
    
    Set Current = Me
    
    If (Comparer Is Nothing) Then
        
        Set Pivot = Element 'try cast
        With Pivot
            Do
                If .IsLessThan(Current.Element) Then
                    Set BestSoFar = Current
                    Set Current = Current.Left
                Else
                    Set Current = Current.Right
                End If
            Loop While Not (Current Is Nothing)
        End With
        
    Else
    
        With Comparer
            Do
                If .AscendingOrdered(Element, Current.Element) Then
                    Set BestSoFar = Current
                    Set Current = Current.Left
                Else
                    Set Current = Current.Right
                End If
            Loop While Not (Current Is Nothing)
        End With
    
    End If
    
    Set GetAfter = BestSoFar
    
End Function

Public Function GetBefore(ByVal Element As Object, Optional ByVal Comparer As IGenericComparer = Nothing) As GenericTreeNode
    
    Dim Current As GenericTreeNode
    Dim BestSoFar As GenericTreeNode
    Dim Pivot As IGenericValue
    
    Set Current = Me
    
    If (Comparer Is Nothing) Then
    
        Set Pivot = Element 'try cast
        With Pivot
            Do
                If .IsGreaterThan(Current.Element) Then
                    Set BestSoFar = Current
                    Set Current = Current.Right
                Else
                    Set Current = Current.Left
                End If
            Loop While Not (Current Is Nothing)
        End With
        
    Else
    
        With Comparer
            Do
                If .AscendingOrdered(Current.Element, Element) Then
                    Set BestSoFar = Current
                    Set Current = Current.Right
                Else
                    Set Current = Current.Left
                End If
            Loop While Not (Current Is Nothing)
        End With
        
    End If
    
    Set GetBefore = BestSoFar
    
End Function

Public Function GetRankOf(ByVal Element As Object, Optional ByVal Comparer As IGenericComparer = Nothing) As Long
    
    Dim Current As GenericTreeNode, Left As GenericTreeNode
    Dim Pivot As IGenericValue

    Set Current = Me
    
    If (Comparer Is Nothing) Then
        
        Set Pivot = Element 'try cast
        With Pivot
            Do
            
                Set Left = Current.Left
                If .Equals(Current.Element) Then
                    If Not (Left Is Nothing) Then _
                        GetRankOf = GetRankOf + Left.Count
                    Exit Function
                End If
                    
                If .IsLessThan(Current.Element) Then
                    Set Current = Left
                    
                Else
                    
                    GetRankOf = GetRankOf + 1
                    If Not (Left Is Nothing) Then _
                        GetRankOf = GetRankOf + Left.Count
                        
                    Set Current = Current.Right
                    
                End If
                
            Loop While Not (Current Is Nothing)
        End With
        
    Else
    
        With Comparer
            Do
            
                Set Left = Current.Left
                If .Equality(Element, Current.Element) Then
                    If Not (Left Is Nothing) Then _
                        GetRankOf = GetRankOf + Left.Count
                    Exit Function
                End If
                    
                If .AscendingOrdered(Element, Current.Element) Then
                    Set Current = Left
                    
                Else
                    
                    GetRankOf = GetRankOf + 1
                    If Not (Left Is Nothing) Then _
                        GetRankOf = GetRankOf + Left.Count
                        
                    Set Current = Current.Right
                    
                End If
                
            Loop While Not (Current Is Nothing)
        End With
        
    End If

End Function

Public Function GetIndexOf(ByVal Element As Object, Optional ByVal Comparer As IGenericComparer = Nothing) As Long
    
    Dim Current As GenericTreeNode, Left As GenericTreeNode
    Dim Pivot As IGenericValue
    
    Set Current = Me
    
    If (Comparer Is Nothing) Then
        
        Set Pivot = Element 'try cast
        With Pivot
            Do
            
                Set Left = Current.Left
                If .Equals(Current.Element) Then
                    If Not (Left Is Nothing) Then _
                        GetIndexOf = GetIndexOf + Left.Count
                    Exit Function
                End If
                    
                If .IsLessThan(Current.Element) Then
                    Set Current = Left
                Else
                    
                    GetIndexOf = GetIndexOf + 1
                    If Not (Left Is Nothing) Then _
                        GetIndexOf = GetIndexOf + Left.Count
                        
                    Set Current = Current.Right
                    
                End If
                
            Loop While Not (Current Is Nothing)
        End With
    
    Else
    
        With Comparer
            Do
              
                Set Left = Current.Left
                If .Equality(Element, Current.Element) Then
                    If Not (Left Is Nothing) Then _
                        GetIndexOf = GetIndexOf + Left.Count
                    Exit Function
                End If
                    
                If .AscendingOrdered(Element, Current.Element) Then
                    Set Current = Left
                Else
                    
                    GetIndexOf = GetIndexOf + 1
                    If Not (Left Is Nothing) Then _
                        GetIndexOf = GetIndexOf + Left.Count
                        
                    Set Current = Current.Right
                    
                End If
                
            Loop While Not (Current Is Nothing)
        End With
        
    End If
    
    GetIndexOf = -1
        
End Function

Public Function ContainsAll(ByVal Sequence As IGenericSequence, Optional ByVal Comparer As IGenericComparer = Nothing) As Boolean

    '@Ignore VariableNotAssigned
    Dim Candidate As Object
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function ContainsAll").IsNothing("Sequence")
    
    If Not (Sequence Is Me) Then
        With Sequence.Iterator
            
            If (Comparer Is Nothing) Then
        
                Do While .HasNext(Candidate)
                    If (Me.Find(Candidate) Is Nothing) Then _
                        Exit Function
                Loop
                
            Else
                
                Do While .HasNext(Candidate)
                    If (Me.FindByComparer(Candidate, Comparer) Is Nothing) Then _
                        Exit Function
                Loop
            
            End If
     
        End With
    End If
    
    ContainsAll = True
    
End Function

Public Function ToArray() As GenericArray
    
    Dim Current As GenericTreeNode
    Dim Elements As GenericArray
    Dim i As Long
    
    Set ToArray = GenericArray
    
    If (Count = 0) Then _
        Exit Function
        
    Set Elements = GenericArray.Build(Count)
    
    Set Current = Me.GetMin
        
    i = Elements.LowerBound - 1
    Do
        i = i + 1
        Set Elements.ElementAt(i) = Current.Element
        Set Current = Current.After
    Loop While Not (Current Is Nothing)

    Set ToArray = Elements
    
End Function

Public Sub CopyTo( _
                    ByVal Target As GenericArray, _
                    ByVal Index As Long, _
                    Optional ByVal BeginAt As Long = 0, _
                    Optional ByVal Elements As Long = 0)
    
    Dim Current As GenericTreeNode
    Dim i As Long
    
    If (Target Is Nothing) Then _
        Call GenericError.Build(Me, "Public Sub CopyTo").IsNothing("Target")
    
    If (Target.IsEmpty) Then _
        Call GenericError.Build(Me, "Public Sub CopyTo").IsNothing("(Target.IsEmpty)")
        
    If (Elements <= 0) Then _
        Elements = Count - BeginAt
        
    If (Count - BeginAt < Elements) Then _
        Call GenericError.Build(Me, "Public Sub CopyTo").ArgOutOfRange("(Count - BeginAt < Elements)")
    
    If (BeginAt <= 0) Then
        Set Current = Me.GetMin
    Else
        Set Current = Me.GetAt(BeginAt)
    End If
    
    Do While (i < Elements)
         With Target
            Set .ElementAt(Index) = Current.Element
            Index = Index + 1
            i = i + 1
            Set Current = Current.After
        End With
    Loop
    
End Sub

Public Function EqualsTo(ByVal Sequence As IGenericSequence, Optional ByVal Comparer As IGenericComparer = Nothing) As Boolean
    
    '@Ignore VariableNotAssigned
    Dim Candidate As Object, Element As IGenericValue
    Dim Current As GenericTreeNode
    
    If (Sequence Is Nothing) Then _
        Exit Function
    
    If (Me Is Sequence) Then
        EqualsTo = True
        Exit Function
    End If
    
    Set Current = Me
        
    With Sequence.Iterator
    
        If Not .HasNext(Candidate) Then _
            Exit Function
            
        If (Comparer Is Nothing) Then
        
            Do
            
                Set Element = Current.Element 'Try cast
                If Not Element.Equals(Candidate) Then _
                    Exit Function
                
                Set Current = Current.After
                If (Current Is Nothing) Then _
                    Exit Function
                    
            Loop While .HasNext(Candidate)
        
        Else

            Do
            
                If Not Comparer.Equality(Current.Element, Candidate) Then _
                    Exit Function
                
                Set Current = Current.After
                If (Current Is Nothing) Then _
                    Exit Function
                    
            Loop While .HasNext(Candidate)

        End If
        
        EqualsTo = Not .HasNext(Candidate)
        
    End With
    
End Function

Public Sub Clear()

    Dim Left As GenericTreeNode
    Dim Right As GenericTreeNode
    Dim Current As GenericTreeNode
    
    Set Current = Me
    
    Do
        With Current
        
            Set .Parent = Nothing
            Set Left = .Left
            
            If Not (Left Is Nothing) Then
                Set .Left = Left.Right 'Rotate the left child up.
                Set Left.Right = Current
                Set Current = Left
            Else
                Set Right = .Right
                Set .Right = Nothing
                Set Current = Right
            End If
        End With
        
    Loop While Not (Current Is Nothing)
    
    Count = 0
    Set Element = Nothing
    
End Sub


