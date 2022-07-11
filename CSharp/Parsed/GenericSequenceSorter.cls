VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSequenceSorter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@IgnoreModule SuspiciousPredeclaredInstanceAccess, AssignedByValParameter, UnassignedVariableUsage, VariableNotAssigned
'@Folder "Sequences"
Option Explicit

Implements IGenericSequenceProvider
Implements IGenericSequence
Implements IGenericIterator

Private Const CACHE_START_SIZE As Long = 2 ^ 13

Private Type Buffer
    Cache As GenericArray
    Index As Long
    Count As Long
End Type

Private Type Member
    Source As IGenericSequence
    AscendingOrdered As IGenericPredicator
    Mapper As GenericMapper
    Direction As SortOrder
    Skip As Long
    Limit As Long
End Type
Private This As Member

Public Property Get Source() As IGenericSequence: Set Source = This.Source: End Property
Public Property Get Skip() As Long: Skip = This.Skip: End Property
Public Property Get Limit() As Long: Limit = This.Limit: End Property
Public Property Get Direction() As SortOrder: Direction = This.Direction: End Property
Public Property Get Mapper() As GenericMapper: Set Mapper = This.Mapper: End Property
Public Property Get AscendingOrdered() As IGenericPredicator: Set AscendingOrdered = This.AscendingOrdered: End Property

Private Function IGenericSequence_Iterator() As IGenericIterator
    
    Dim Buffer As Buffer
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequence_Iterator").IsNothing("Sequence")
    
    Set IGenericSequence_Iterator = Me 'avoid null
    
    If Not (This.Limit = 1) Then
    
        If Not Cache(This.Source, Buffer) Then _
            Exit Function
        
        If Not Sort(Buffer, This.Skip, This.Limit, This.Mapper, This.AscendingOrdered) Then _
            Exit Function
        
    Else 'edge case
    
        Set Buffer.Cache = GenericArray.Build(1)
        Set Buffer.Cache.ElementAt(GenericArray.LowerBound) = IGenericSequenceProvider_First
        Buffer.Count = 1
    
    End If
    
    If (This.Skip > 0) Then _
        Buffer.Index = This.Skip
    
    If (This.Limit > 0) Then
        If (This.Limit > Buffer.Count - Buffer.Index) Then
            Buffer.Count = Buffer.Count - Buffer.Index
        Else
            Buffer.Count = This.Limit
        End If
    End If
        
    Set IGenericSequence_Iterator = Buffer.Cache.Range(Buffer.Index, Buffer.Count)
    
End Function

'@Ignore ProcedureNotUsed
Public Function Of( _
                        ByVal Source As IGenericSequence, _
                        ByVal Direction As SortOrder, _
                        Optional ByVal Mapper As GenericMapper, _
                        Optional ByVal AscendingOrdered As IGenericPredicator, _
                        Optional ByVal Skip As Long = 0, _
                        Optional ByVal Limit As Long = 0) As IGenericSequence

    If (Source Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Of").IsNothing("Sequence")
    
    If (Skip < 0) Then _
        Call GenericError.Build(Me, "Public Function Of").InvalidInput("(Skip < 0)")
        
    If (Limit < 0) Then _
        Call GenericError.Build(Me, "Public Function Of").InvalidInput("(Limit < 0)")
    
    If Not Me Is GenericSequenceSorter Then _
        Call GenericError.Build(Me, "Public Function Of").InvalidOperation("instance call")
        
    This.Skip = Skip
    This.Limit = Limit
    This.Direction = Direction
    Set This.Source = Source
    Set This.Mapper = Mapper
    Set This.AscendingOrdered = AscendingOrdered
    
    Set Of = New GenericSequenceSorter
    This = EmptyStructure
    
End Function
'@Ignore NonReturningFunction, EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Sub Class_Initialize()
  
    With GenericSequenceSorter
        This.Direction = .Direction
        This.Skip = .Skip
        This.Limit = .Limit
        Set This.AscendingOrdered = .AscendingOrdered
        Set This.Mapper = .Mapper
        Set This.Source = .Source
    End With
    
End Sub

Private Function Cache(ByVal Source As IGenericSequence, ByRef Buffer As Buffer) As Boolean

    Dim tmp As GenericArray
    Dim Element As Object
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    Dim Capacity As Long
  
    If (Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Sub Cache").IsNothing("Sequence")
    
    'avoid null
    If (Buffer.Cache Is Nothing) Then
        Set Buffer.Cache = GenericArray
        Buffer.Index = Buffer.Cache.LowerBound
    End If
        
    If TypeOf Source Is IGenericCollection Then
    
        '@Ignore SetAssignmentWithIncompatibleObjectType
        Set Collection = Source
        Capacity = Collection.Count + Buffer.Count

        If (Capacity > 0) Then
        
            If (Buffer.Cache.Length < Capacity) Then
                Set tmp = GenericArray.Build(Capacity)
                'In case we decide to allow an "Add-Method" in the future, previous cached elements will be copied
                If (Buffer.Count > 0) Then _
                    Call Buffer.Cache.CopyTo(Target:=tmp, Index:=tmp.LowerBound, BeginAt:=Buffer.Cache.LowerBound, Count:=Buffer.Count)
                    
                Set Buffer.Cache = tmp

            End If
            
            Call Collection.CopyTo(Buffer.Cache, Buffer.Count)
            Buffer.Count = Capacity
            
        End If

    Else
        
        With Source.Iterator 'Compute cache via the hard way
            If .HasNext(Element) Then
                    
                If TypeOf This.Source Is IGenericSequenceProvider Then
                    Set Provider = This.Source
                    Capacity = Provider.GuessCount + Buffer.Count
                End If
        
                If (Capacity > 0) Then
                    If (Buffer.Cache.Length < Capacity) Then
                        Set tmp = GenericArray.Build(Capacity)
                        'In case we decide to allow an "Add-Method" in the future, previous cached elements will be copied
                        If (Buffer.Count > 0) Then _
                            Call Buffer.Cache.CopyTo(Target:=tmp, Index:=tmp.LowerBound, BeginAt:=Buffer.Cache.LowerBound, Count:=Buffer.Count)
                    
                        Set Buffer.Cache = tmp

                    End If
                End If
                
                Capacity = Buffer.Cache.Length
                
                Do
                    If (Buffer.Count = Capacity) Then _
                        Capacity = Expand(Buffer)
                    
                    Set Buffer.Cache.ElementAt(Buffer.Count) = Element
                    Buffer.Count = Buffer.Count + 1
                    
                Loop While .HasNext(Element)
                
            End If
        End With
    
    End If
    
    Cache = Buffer.Count > 0
    
End Function

Private Function Expand(ByRef Buffer As Buffer) As Long
    
    Dim newCapacity As Long
    Dim tmp As GenericArray
    
    With Buffer.Cache
        
        newCapacity = VBA.IIf(.Length = 0, CACHE_START_SIZE, .Length * 2)
        
        If (newCapacity > .MaxLength) Then _
            newCapacity = .MaxLength
                
        Set tmp = GenericArray.Build(newCapacity)
        
        If (Buffer.Count > 0) Then _
            Call .CopyTo(Target:=tmp, Index:=tmp.LowerBound, BeginAt:=.LowerBound, Count:=Buffer.Count)
        
    End With
    
    Set Buffer.Cache = tmp
    
    Expand = newCapacity
    
End Function

Private Function CacheKeys(ByRef Buffer As Buffer, ByVal Mapper As GenericMapper) As GenericArray

    Dim i As Long

    Set CacheKeys = GenericArray.Build(Buffer.Count)

    With Buffer.Cache
        For i = .LowerBound To Buffer.Count - 1: Set CacheKeys.ElementAt(i) = Mapper.Map(.ElementAt(i)): Next
    End With
  
End Function

Private Function FindTopElement( _
                                    ByVal Direction As SortOrder, _
                                    ByVal Sequence As IGenericSequence, _
                                    Optional ByVal Mapper As GenericMapper, _
                                    Optional ByVal AscendingOrdered As IGenericPredicator) As Object
    
    Dim Candidate As Object, Value As IGenericValue, MappedCandidate As Object
    Dim BestSoFar As Object, BestValueSoFar As IGenericValue, MappedBestSoFar As Object
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "FindTopElement").IsNothing("Sequence")
    
    With Sequence.Iterator
        If (Mapper Is Nothing) Then
            If (AscendingOrdered Is Nothing) Then
    
                If .HasNext(BestValueSoFar) Then 'Try convert
                    If (Direction = Ascending) Then
                
                        Do While .HasNext(Value) 'Try convert
                            If Not (Value Is Nothing) Then _
                                If Value.IsLessThan(BestValueSoFar) Then _
                                    Set BestValueSoFar = Value
                        Loop
                        
                    Else
                    
                        Do While .HasNext(Value) 'Try convert
                            If Not (Value Is Nothing) Then _
                                If Value.IsGreaterThan(BestValueSoFar) Then _
                                    Set BestValueSoFar = Value
                        Loop
                    
                    End If
                End If
                
                Set BestSoFar = BestValueSoFar ' avoid converting during search
                
            Else
                
                If .HasNext(BestSoFar) Then
                    If (Direction = Ascending) Then
                    
                        Do While .HasNext(Candidate)
                            If AscendingOrdered.IsTrue(Candidate, BestSoFar) Then _
                                Set BestSoFar = Candidate
                        Loop
                        
                    Else
                        
                        Do While .HasNext(Candidate)
                            If AscendingOrdered.IsTrue(BestSoFar, Candidate) Then _
                                Set BestSoFar = Candidate
                        Loop
                    
                    End If
                End If
                
            End If
            
        Else
    
            If (AscendingOrdered Is Nothing) Then
                If .HasNext(BestSoFar) Then
                    
                    Set BestValueSoFar = Mapper.Map(BestSoFar) 'Try convert
                    
                    If (Direction = Ascending) Then
                    
                       Do While .HasNext(Candidate)
                            Set Value = Mapper.Map(Candidate) 'Try convert
                            If Not (Value Is Nothing) Then
                                If Value.IsLessThan(BestValueSoFar) Then
                                    Set BestSoFar = Candidate
                                    Set BestValueSoFar = Value
                                End If
                            End If
                        Loop
                    
                    Else
                        
                         Do While .HasNext(Candidate)
                            Set Value = Mapper.Map(Candidate) 'Try convert
                            If Not (Value Is Nothing) Then
                                If Value.IsGreaterThan(BestValueSoFar) Then
                                    Set BestSoFar = Candidate
                                    Set BestValueSoFar = Value
                                End If
                            End If
                        Loop
                        
                    End If
                
                End If
                 
            Else
            
                If .HasNext(BestSoFar) Then
                    
                    Set MappedBestSoFar = Mapper.Map(BestSoFar)
                    
                    If (Direction = Ascending) Then
                        
                        Do While .HasNext(Candidate)
                            Set MappedCandidate = Mapper.Map(Candidate)
                            If AscendingOrdered.IsTrue(MappedCandidate, MappedBestSoFar) Then
                                Set BestSoFar = Candidate
                                Set MappedBestSoFar = MappedCandidate
                            End If
                        Loop
                        
                    Else
                        
                        Do While .HasNext(Candidate)
                            Set MappedCandidate = Mapper.Map(Candidate)
                            If AscendingOrdered.IsTrue(MappedBestSoFar, MappedCandidate) Then
                                Set BestSoFar = Candidate
                                Set MappedBestSoFar = MappedCandidate
                            End If
                        Loop
                        
                    End If
                    
                End If
                
            End If
        
        End If
    End With
    
    Set FindTopElement = BestSoFar
    
End Function

Private Function Sort( _
                        ByRef Buffer As Buffer, _
                        Optional ByVal Skip As Long = 0, _
                        Optional ByVal Limit As Long = 0, _
                        Optional ByVal Mapper As GenericMapper, _
                        Optional ByVal AscendingOrdered As IGenericPredicator) As Boolean
    
    If (Skip >= Buffer.Count) Or (Buffer.Count < 2) Then _
        Exit Function
    
    If (Skip = 0 And Limit = 0) Then
    
        If (This.Mapper Is Nothing) Then
            Call Buffer.Cache.Sort(Direction:=This.Direction, AscendingOrdered:=AscendingOrdered, Length:=Buffer.Count)
        Else
            Call CacheKeys(Buffer, Mapper).SortWith(Items:=Buffer.Cache, Direction:=This.Direction, AscendingOrdered:=AscendingOrdered, Length:=Buffer.Count)
        End If
       
    ElseIf (Skip = 0 And Limit > 0) Then
        
        If (This.Mapper Is Nothing) Then
            Call Buffer.Cache.SortTop(k:=This.Limit, Direction:=This.Direction, AscendingOrdered:=AscendingOrdered, Length:=Buffer.Count)
        Else
            Call CacheKeys(Buffer, Mapper).SortTopWith(k:=This.Limit, Items:=Buffer.Cache, Direction:=This.Direction, AscendingOrdered:=AscendingOrdered, Length:=Buffer.Count)
        End If
        
    Else
        
        If (Mapper Is Nothing) Then
            Call Buffer.Cache.SortPartial(BeginAt:=VBA.IIf(Skip <= Buffer.Cache.LowerBound, Buffer.Cache.LowerBound, Skip), EndAt:=VBA.IIf(Limit <= 0, Buffer.Count, Skip + Limit), Direction:=This.Direction, AscendingOrdered:=AscendingOrdered, Length:=Buffer.Count)
        Else
            Call CacheKeys(Buffer, Mapper).SortPartialWith(Items:=Buffer.Cache, BeginAt:=VBA.IIf(Skip = 0, Buffer.Cache.LowerBound, Skip), EndAt:=VBA.IIf(Limit = 0, Buffer.Count, Skip + Limit), Direction:=This.Direction, AscendingOrdered:=AscendingOrdered, Length:=Buffer.Count)
        End If

    End If
    
    Sort = True
    
End Function

Private Function IGenericSequenceProvider_GetAt(ByVal Index As Long) As Object
    
    Dim Buffer As Buffer
    
    If (Index = GenericArray.LowerBound And This.Skip = 0) Then
        Set IGenericSequenceProvider_GetAt = FindTopElement(Direction:=VBA.IIf(This.Direction = Ascending, Descending, Ascending), Sequence:=This.Source, Mapper:=This.Mapper, AscendingOrdered:=This.AscendingOrdered)
        Exit Function
    End If
   
    If (This.Limit > 0) Then _
        If (Index >= This.Limit) Then _
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GetAt").IndexOutofRange("(Index >= This.Limit)")
    
    If Not Cache(This.Source, Buffer) Then _
        Exit Function

    Index = Index + This.Skip
        
    If (Index >= Buffer.Count) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GetAt").IndexOutofRange("(Index >= Buffer.Count)")
            
    If (This.Mapper Is Nothing) Then
        Index = Buffer.Cache.SelectKth(Kth:=Index, Direction:=This.Direction, AscendingOrdered:=This.AscendingOrdered, Length:=Buffer.Count)
    Else
        Index = CacheKeys(Buffer, This.Mapper).SelectKth(Kth:=Index, Direction:=This.Direction, AscendingOrdered:=This.AscendingOrdered, Length:=Buffer.Count)
    End If
        
    Set IGenericSequenceProvider_GetAt = Buffer.Cache.ElementAt(Index)
    
End Function

Private Function IGenericSequenceProvider_Count() As Long
    
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    Dim Count As Long
    
    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "IGenericSequenceProvider_Count").IsNothing("Sequence")
        
        Case (This.Limit > 0)
            With This.Source.Iterator
                Do While .HasNext()
                    Count = Count + 1
                    If (Count = Limit) Then _
                        Exit Do
                Loop
            End With
            
        Case (TypeOf This.Source Is IGenericCollection)
            '@Ignore SetAssignmentWithIncompatibleObjectType
            Set Collection = This.Source
            Count = Collection.Count
        
        Case (TypeOf This.Source Is IGenericSequenceProvider)
            Set Provider = This.Source
            Count = Provider.Count
           
        Case Else
            With This.Source.Iterator
                Do While .HasNext(): Count = Count + 1: Loop
            End With
            
    End Select
  
    IGenericSequenceProvider_Count = Count
    
End Function

Private Function IGenericSequenceProvider_GuessCount() As Long
    
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    Dim guess As Long
    
    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GuessCount").IsNothing("Sequence")
            
        Case (TypeOf This.Source Is IGenericCollection)
            '@Ignore SetAssignmentWithIncompatibleObjectType
            Set Collection = This.Source
            guess = Collection.Count
        
        Case (TypeOf This.Source Is IGenericSequenceProvider)
            Set Provider = This.Source
            guess = Provider.GuessCount
           
        Case Else
            guess = -1
            
    End Select
    
    IGenericSequenceProvider_GuessCount = VBA.IIf(guess > This.Limit, This.Limit, guess)
    
End Function

Private Function IGenericSequenceProvider_First() As Object
    
    Dim Buffer As Buffer
    Dim Index As Long
    
    If (This.Skip = 0) Then
        
        Set IGenericSequenceProvider_First = FindTopElement(Direction:=This.Direction, Sequence:=This.Source, Mapper:=This.Mapper, AscendingOrdered:=This.AscendingOrdered)
   
    ElseIf Cache(This.Source, Buffer) Then
         
        If (This.Mapper Is Nothing) Then
            Index = Buffer.Cache.SelectKth(Kth:=This.Skip, Direction:=This.Direction, AscendingOrdered:=This.AscendingOrdered, Length:=Buffer.Count)
        Else
            Index = CacheKeys(Buffer, This.Mapper).SelectKth(Kth:=This.Skip, Direction:=This.Direction, AscendingOrdered:=This.AscendingOrdered, Length:=Buffer.Count)
        End If
            
        Set IGenericSequenceProvider_First = Buffer.Cache.ElementAt(Index)
    
    End If
    
End Function

Private Function IGenericSequenceProvider_Last() As Object
    
    Dim Buffer As Buffer
    Dim Index As Long
    
    If (This.Limit = 0) Then
        
        Set IGenericSequenceProvider_Last = FindTopElement(Direction:=This.Direction, Sequence:=This.Source, Mapper:=This.Mapper, AscendingOrdered:=This.AscendingOrdered)
   
    ElseIf Cache(This.Source, Buffer) Then
         
        If (This.Mapper Is Nothing) Then
            Index = Buffer.Cache.SelectKth(Kth:=This.Skip + This.Limit - 1, Direction:=This.Direction, AscendingOrdered:=This.AscendingOrdered, Length:=Buffer.Count)
        Else
            Index = CacheKeys(Buffer, This.Mapper).SelectKth(Kth:=This.Skip + This.Limit - 1, Direction:=This.Direction, AscendingOrdered:=This.AscendingOrdered, Length:=Buffer.Count)
        End If
            
        Set IGenericSequenceProvider_Last = Buffer.Cache.ElementAt(Index)
    
    End If
    
End Function

Private Function IGenericSequenceProvider_ToArray() As GenericArray
    
    Dim Buffer As Buffer
    Dim Provider As IGenericSequenceProvider
    
    Dim BeginAt As Long, EndAt As Long
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_ToArray").IsNothing("Sequence")
        
    If (This.Limit = 1) Then  'edge case
        Set IGenericSequenceProvider_ToArray = GenericArray.Build(1)
        Set IGenericSequenceProvider_ToArray.ElementAt(GenericArray.LowerBound) = IGenericSequenceProvider_First
        Exit Function
    End If
    
    Set IGenericSequenceProvider_ToArray = GenericArray 'avoid null
    
    If (TypeOf This.Source Is IGenericSequenceProvider) Then 'fast path
    
        Set Provider = This.Source
        Set Buffer.Cache = Provider.ToArray
        Buffer.Count = Buffer.Cache.Length
    
    ElseIf Not Cache(This.Source, Buffer) Then 'early out
        Exit Function
    
    End If
    'early out
    If Not Sort(Buffer, This.Skip, This.Limit, This.Mapper, This.AscendingOrdered) Then _
        Exit Function
    
    If (This.Skip > 0) Then
        BeginAt = This.Skip
    Else
        BeginAt = Buffer.Cache.LowerBound
    End If
    
    EndAt = Buffer.Count - 1
    
    If (This.Limit > 0) Then _
        If (BeginAt + This.Limit <= EndAt) Then _
            EndAt = BeginAt + This.Limit - 1
   
    With Buffer.Cache
    
        If (Buffer.Count = .Length) Then
            If (BeginAt = .LowerBound) And (EndAt = .Length - 1) Then
                Set IGenericSequenceProvider_ToArray = Buffer.Cache
                Exit Function
            End If
        End If
            
        Set IGenericSequenceProvider_ToArray = GenericArray.Build(EndAt - BeginAt + 1)
        Call .CopyTo(Target:=IGenericSequenceProvider_ToArray, Index:=IGenericSequenceProvider_ToArray.LowerBound, BeginAt:=BeginAt, Count:=IGenericSequenceProvider_ToArray.Length)
        
'        For i = BeginAt To EndAt: Set IGenericSequenceProvider_ToArray.ElementAt(i - BeginAt) = .ElementAt(i): Next
    End With
    
End Function

Private Function IGenericSequenceProvider_ToList() As GenericOrderedList
   
    Dim Buffer As Buffer
    Dim Provider As IGenericSequenceProvider
    Dim Result As GenericOrderedList
    
    Dim i As Long, BeginAt As Long, EndAt As Long
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_ToList").IsNothing("Sequence")
        
    If (This.Limit = 1) Then 'edge case
        Set IGenericSequenceProvider_ToList = GenericOrderedList.Build(1)
        Call IGenericSequenceProvider_ToList.Add(IGenericSequenceProvider_First)
        Exit Function
    End If
    
    If (TypeOf This.Source Is IGenericSequenceProvider) Then 'fast path
        
        Set Provider = This.Source
        If (This.Mapper Is Nothing And This.Skip = 0 And This.Limit = 0) Then 'we can do even better
            
            Set IGenericSequenceProvider_ToList = Provider.ToList.Sort(Direction:=This.Direction, AscendingOrdered:=This.AscendingOrdered)
            Exit Function
            
        Else
        
            Set Buffer.Cache = Provider.ToArray
            Buffer.Count = Buffer.Cache.Length
    
        End If
        
    ElseIf Not Cache(This.Source, Buffer) Then  'early out
        
        Set IGenericSequenceProvider_ToList = GenericOrderedList.Build
        Exit Function
    
    End If
    
    Set Result = GenericOrderedList.Build
    Set IGenericSequenceProvider_ToList = Result
    
    'early out
    If Not Sort(Buffer, This.Skip, This.Limit, This.Mapper, This.AscendingOrdered) Then _
        Exit Function
        
    If (This.Skip > 0) Then
        BeginAt = This.Skip
    Else
        BeginAt = Buffer.Cache.LowerBound
    End If
    
    EndAt = Buffer.Count - 1
    
    If (This.Limit > 0) Then _
        If (BeginAt + This.Limit <= EndAt) Then _
            EndAt = BeginAt + This.Limit - 1
    
    With Buffer.Cache
    
        If (Buffer.Count = .Length) Then
            If (BeginAt = .LowerBound) And (EndAt = .Length - 1) Then
                Set IGenericSequenceProvider_ToList = Result.AddAll(Buffer.Cache)
                Exit Function
            End If
        End If
         
        Call Result.EnsureCapacity(EndAt - BeginAt + 1)
        For i = BeginAt To EndAt: Call Result.Add(.ElementAt(i)): Next
    
    End With
    
    Set IGenericSequenceProvider_ToList = Result
    
End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As Object: Set IGenericIterator_Current = Nothing: End Property
Private Function IGenericIterator_HasNext(Optional Out As Object) As Boolean: IGenericIterator_HasNext = False: End Function


