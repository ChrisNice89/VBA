VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder "Sequences"
'@IgnoreModule VariableNotAssigned, SuspiciousPredeclaredInstanceAccess, AssignedByValParameter, UnassignedVariableUsage, ProcedureNotUsed
Option Explicit

'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericSequence
'*
'* Purpose: The methods in this class provide an implementation of the standard query operators for querying data that implement IGenericSequence.
'*          The standard query operators are general purpose methods and enable you to express traversal, filter, and projection operations over data.
'*          The majority of the methods in this class can be called like an instance method on any object that implements IGenericSequence.
'*          Methods that are used in a query that returns a sequence of values do not consume the target data until the query object is iterated.
'*          This is known as deferred execution. Methods that are used in a query that returns a singleton value ( e.g. Sum, Count ect.) execute and consume the target data immediately.
'*
'*
'* Notes: Provides powerful filtering, ordering (and grouping) capabilities with a minimum of application code.
'*        Optimized way to perform set based operations on in memory objects
'*        In general, the more complex the operation you want to perform on the data, the more benefit you'll realize by using GenericSequence instead of traditional iteration techniques.
'*        ReadOnly access
'*
'*
'*======================================================================================================================

Implements IGenericSequence
Implements IGenericHashProvider
Implements IGenericPredicator

Private Type Member
    Sequence As IGenericSequence
End Type
Private This As Member

Public Property Get SequenceEquality() As IGenericPredicator: Set SequenceEquality = GenericSequence: End Property
Public Property Get Source() As IGenericSequence: Set Source = This.Sequence: End Property

Public Property Get Iterator() As IGenericIterator

    If Not (This.Sequence Is Nothing) Then
        Set Iterator = This.Sequence.Iterator
    Else
        Set Iterator = IGenericSequenceProvider
    End If
    
End Property

Private Function IGenericSequence_Iterator() As IGenericIterator
   
    If (This.Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequence_Iterator").IsNothing("Sequence")
    
    Set IGenericSequence_Iterator = This.Sequence.Iterator
    
End Function
'Extension Method
Public Function Of(ParamArray IGenericValues() As Variant) As GenericSequence
    
    Dim i As Long
    Dim Values As GenericArray
    
    If VBA.IsMissing(IGenericValues) Then _
        Call GenericError.Build(Me, "Public Function Of").IsNothing("ParamArray IGenericValues are missing")
    
    Set Values = GenericArray.Build(UBound(IGenericValues) + 1)
    
    With Values
        For i = LBound(IGenericValues) To UBound(IGenericValues)
            
            If Not VBA.IsObject(IGenericValues(i)) Then _
                Call GenericError.Build(Me, "Public Function Of").InvalidInput("Datatype {" & VBA.TypeName$(IGenericValues(i)) & "} not allowed")
            
            If (IGenericValues(i) Is Nothing) Then _
                Call GenericError.Build(Me, "Public Function Of").IsNothing("IGenericValues(i)")
            
            If Not TypeOf IGenericValues(i) Is IGeneric Then _
                Call GenericError.Build(Me, "Public Function Of").InvalidInput("Not TypeOf IGenericValues(i) Is IGeneric")
                
            Set .ElementAt(i) = IGenericValues(i)
          
        Next
    End With
    
    Set Of = GenericSequence.Stream(Values)
    
End Function

Public Function Stream(ByVal Sequence As IGenericSequence) As GenericSequence

    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Stream").IsNothing("Sequence")
    
    If Not Me Is GenericSequence Then _
        Call GenericError.Build(Me, "Public Function Of").InvalidOperation("instance call")
         
    If TypeOf Sequence Is GenericSequence Then _
        Set Sequence = GetSource(Sequence)
    
    Set This.Sequence = Sequence
    
    Set Stream = New GenericSequence
    This = EmptyStructure
    
End Function

'@Ignore NonReturningFunction, EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Sub Class_Initialize(): Set This.Sequence = GenericSequence.Source: End Sub

Public Function Apply(ByVal Mapper As GenericMapper) As GenericSequence
    
    If (This.Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Apply").IsNothing("Sequence")
    
    If (Mapper Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Apply").IsNothing("Mapper")
    
    Set This.Sequence = GenericSequenceMapper.Of(Source:=This.Sequence, Mapper:=Mapper)

    Set Apply = Me
    
End Function

Public Function Where(ByVal Predicate As IGenericPredicate) As GenericSequence
    
    If (This.Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Where").IsNothing("Sequence")
    
    If (Predicate Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Where").IsNothing("Predicate")
    
    Select Case True
        Case TypeOf This.Sequence Is GenericSequenceSorter
            With SequenceSorter(This.Sequence) ' elements that are not needed will be not sorted
                Set This.Sequence = GenericSequenceSorter.Of(Source:=GenericSequenceWhere.Of(Source:=.Source, Predicate:=Predicate), Direction:=.Direction, Mapper:=.Mapper, AscendingOrdered:=.AscendingOrdered, Skip:=.Skip, Limit:=.Limit)
            End With
            
        Case Else
            Set This.Sequence = GenericSequenceWhere.Of(Source:=This.Sequence, Predicate:=Predicate)
            
    End Select
    
    Set Where = Me
    
End Function

Public Function WhereIn( _
                            ByVal Other As IGenericSequence, _
                            Optional ByVal Mapper As GenericMapper, _
                            Optional ByVal HashProvider As IGenericHashProvider) As GenericSequence
    
    If (This.Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function WhereIn").IsNothing("Sequence")
    
    If (Other Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function WhereIn").IsNothing("Other")

    If TypeOf Other Is GenericSequence Then _
        Set Other = GetSource(Other)
    
    Select Case True
        Case TypeOf This.Sequence Is GenericSequenceSorter
            With SequenceSorter(This.Sequence) ' elements that are not needed will be sorted
                Set This.Sequence = GenericSequenceSorter.Of(Source:=GenericSequenceSet.Of(Source:=.Source, Other:=Other, Mapper:=Mapper, HashProvider:=HashProvider, Complement:=False), Direction:=.Direction, Mapper:=.Mapper, AscendingOrdered:=.AscendingOrdered, Skip:=.Skip, Limit:=.Limit)
            End With
        
        Case TypeOf This.Sequence Is GenericSequenceDistinct
            With SequenceDistinct(This.Sequence) 'Switch operations -> less memory alloc
                Set This.Sequence = GenericSequenceDistinct.Of(Source:=GenericSequenceSet.Of(Source:=.Source, Other:=Other, Mapper:=Mapper, HashProvider:=HashProvider, Complement:=False), HashProvider:=.HashProvider)
            End With
        
        Case Else
            Set This.Sequence = GenericSequenceSet.Of(Source:=This.Sequence, Other:=Other, Mapper:=Mapper, HashProvider:=HashProvider, Complement:=False)
        
    End Select
    
    Set WhereIn = Me
    
End Function

Public Function Except( _
                         ByVal Other As IGenericSequence, _
                         Optional ByVal Mapper As GenericMapper, _
                         Optional ByVal HashProvider As IGenericHashProvider) As GenericSequence
                            
    If (This.Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Except").IsNothing("Sequence")
    
    If (Other Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Except").IsNothing("Other")
    
    If TypeOf Other Is GenericSequence Then _
        Set Other = GetSource(Other)
        
    Select Case True
        Case TypeOf This.Sequence Is GenericSequenceSorter
            With SequenceSorter(This.Sequence) 'Except may not keep the order of elements and elements that are not needed will be sorted
                Set This.Sequence = GenericSequenceSorter.Of(Source:=GenericSequenceSet.Of(Source:=.Source, Other:=Other, Mapper:=Mapper, HashProvider:=HashProvider, Complement:=True), Direction:=.Direction, Mapper:=.Mapper, AscendingOrdered:=.AscendingOrdered, Skip:=.Skip, Limit:=.Limit)
            End With
        
        Case TypeOf This.Sequence Is GenericSequenceDistinct
            With SequenceDistinct(This.Sequence) 'Switch operations -> less memory alloc
                Set This.Sequence = GenericSequenceDistinct.Of(Source:=GenericSequenceSet.Of(Source:=.Source, Other:=Other, Mapper:=Mapper, HashProvider:=HashProvider, Complement:=True), HashProvider:=.HashProvider)
            End With
        
        Case Else
            Set This.Sequence = GenericSequenceSet.Of(Source:=This.Sequence, Other:=Other, Mapper:=Mapper, HashProvider:=HashProvider, Complement:=True)
        
    End Select
    
    Set Except = Me
    
End Function

Public Function Distinct(Optional ByVal HashProvider As IGenericHashProvider) As GenericSequence
  
    If (This.Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Distinct").IsNothing("Sequence")
    
    Select Case True
        Case TypeOf This.Sequence Is GenericSequenceSorter
            With SequenceSorter(This.Sequence) 'Distinct may not keep the order of elements and elements that are not needed will be sorted
                Set This.Sequence = GenericSequenceSorter.Of(Source:=GenericSequenceDistinct.Of(Source:=.Source, HashProvider:=HashProvider), Direction:=.Direction, Mapper:=.Mapper, AscendingOrdered:=.AscendingOrdered, Skip:=.Skip, Limit:=.Limit)
            End With
        
        Case Else
            Set This.Sequence = GenericSequenceDistinct.Of(Source:=This.Sequence, HashProvider:=HashProvider)
        
    End Select
    
    Set Distinct = Me
    
End Function

Public Function Partition(ByVal Index As Long, ByVal Size As Long) As GenericSequence
    
    Select Case True
        Case (This.Sequence Is Nothing)
            Call GenericError.Build(Me, "Public Function Partition").IsNothing("Sequence")
        
        Case (Index < 0)
            Call GenericError.Build(Me, "Public Function Partition").ArgOutOfRange("PageIndex < 0")
            
        Case (Size < 0)
            Call GenericError.Build(Me, "Public Function Partition").ArgOutOfRange("PageSize < 0")
            
        Case Else
            If (Index = 0) Then
                Set This.Sequence = Me.Take(Size).Source
            Else
                Set This.Sequence = Me.Skip(Index * Size).Take(Size).Source
            End If
    End Select
    
    Set Partition = Me
    
End Function

Public Function Skip(ByVal Count As Long) As GenericSequence
    
    If (Count > 0) Then
        Select Case True
            Case (This.Sequence Is Nothing)
                Call GenericError.Build(Me, "Public Function Skip").IsNothing("Sequence")
            
            Case TypeOf This.Sequence Is GenericSequenceMapper 'For performance put a skip operation before mapping
                With SequenceMapper(This.Sequence)
                    Set This.Sequence = GenericSequenceMapper.Of(Source:=GenericSequenceSkip.Of(Source:=.Source, Skip:=Count), Mapper:=.Mapper)
                End With
            
            Case TypeOf This.Sequence Is GenericSequenceSorter
                With SequenceSorter(This.Sequence)
                    Set This.Sequence = GenericSequenceSorter.Of(Source:=.Source, Direction:=.Direction, Mapper:=.Mapper, AscendingOrdered:=.AscendingOrdered, Skip:=.Skip + Count, Limit:=.Limit)
                End With
            
            Case TypeOf This.Sequence Is GenericSequenceTake 'For performance put a skip operation before the take
                With SequenceTake(This.Sequence)
                    If (.Take > Count) Then
                        Set This.Sequence = GenericSequenceTake.Of(Source:=GenericSequenceSkip.Of(Source:=.Source, Skip:=Count), Take:=.Take - Count)
                    Else
                        Set This.Sequence = IGenericSequenceProvider 'Empty
                    End If
                End With
                
            Case Else
                Set This.Sequence = GenericSequenceSkip.Of(Source:=This.Sequence, Skip:=Count)
        
        End Select
    End If
    
    Set Skip = Me
    
End Function

Public Function Take(ByVal Count As Long) As GenericSequence
    
    If (This.Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Take").IsNothing("Sequence")
    
    Select Case True
        Case (Count <= 0)
            Set This.Sequence = IGenericSequenceProvider 'Empty
    
        Case TypeOf This.Sequence Is GenericSequenceSorter
            With SequenceSorter(This.Sequence)
                Set This.Sequence = GenericSequenceSorter.Of(Source:=.Source, Direction:=.Direction, Mapper:=.Mapper, AscendingOrdered:=.AscendingOrdered, Skip:=.Skip, Limit:=Count)
            End With
    
        Case Else
            Set This.Sequence = GenericSequenceTake.Of(Source:=This.Sequence, Take:=Count)
            
    End Select
      
    Set Take = Me
    
End Function

Public Function OrderBy( _
                            ByVal Mapper As GenericMapper, _
                            Optional ByVal Direction As SortOrder = SortOrder.Ascending, _
                            Optional ByVal AscendingOrdered As IGenericPredicator) As GenericSequence
    
    If (This.Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function OrderBy").IsNothing("Sequence")
    
    If (Mapper Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function OrderBy").IsNothing("Mapper")
    
    Set This.Sequence = GenericSequenceSorter.Of(Source:=This.Sequence, Direction:=Direction, Mapper:=Mapper, AscendingOrdered:=AscendingOrdered)
    Set OrderBy = Me
   
End Function

Public Function Ascending( _
                            Optional ByVal Mapper As GenericMapper, _
                            Optional ByVal AscendingOrdered As IGenericPredicator) As GenericSequence
    
    If (This.Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Ascending").IsNothing("Sequence")
        
    Set This.Sequence = GenericSequenceSorter.Of(Source:=This.Sequence, Direction:=SortOrder.Ascending, Mapper:=Mapper, AscendingOrdered:=AscendingOrdered)
    Set Ascending = Me
   
End Function

Public Function Descending( _
                            Optional ByVal Mapper As GenericMapper, _
                            Optional ByVal AscendingOrdered As IGenericPredicator) As GenericSequence
    
    If (This.Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Descending").IsNothing("Sequence")
    
    Set This.Sequence = GenericSequenceSorter.Of(Source:=This.Sequence, Direction:=SortOrder.Descending, Mapper:=Mapper, AscendingOrdered:=AscendingOrdered)
    Set Descending = Me
   
End Function

Public Function Append(ByVal Other As IGenericSequence) As GenericSequence
     
    If (This.Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Append").IsNothing("Sequence")
    
    If (Other Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Append").IsNothing("Sequence")
    
    If TypeOf Other Is GenericSequence Then _
        Set Other = GetSource(Other)
    
    Set This.Sequence = GenericSequenceConcat.Of(This.Sequence, Other)
    Set Append = Me
    
End Function

Public Function ForAll(ByVal Predicate As IGenericPredicate) As Boolean

    Dim Element As Object
    
    If (This.Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Sub ForAll").IsNothing("Sequence")
    
    If (Predicate Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function ForAll").IsNothing("(Delegate Is Nothing)")
  
    With This.Sequence.Iterator
        
        If Not .HasNext(Element) Then _
            Call GenericError.Build(Me, "Public Function ForAll").InvalidOperation
            
        Do
            If Not Predicate.Match(Element) Then _
                Exit Function
        Loop While .HasNext(Element)
        
    End With
    
    ForAll = True
    
End Function

Public Function HasAny() As Boolean
    
    Select Case True
        Case (This.Sequence Is Nothing)
            Call GenericError.Build(Me, "Public Function HasAny").IsNothing("Sequence")
        
        Case TypeOf This.Sequence Is IGenericCollection
            '@Ignore ArgumentWithIncompatibleObjectType
            HasAny = CastCollection(This.Sequence).Count > 0
            
        Case Else
            HasAny = This.Sequence.Iterator.HasNext
            
    End Select
    
End Function

Public Function Contains(ByVal Element As Object, Optional ByVal Equality As IGenericPredicator) As Boolean
    
    Dim Candidate As Object
    Dim Casted As IGenericValue
    
    Select Case True
        Case (This.Sequence Is Nothing)
            Call GenericError.Build(Me, "Public Function Contains").IsNothing("Sequence")
        
        Case Not (Equality Is Nothing)
            With This.Sequence.Iterator
                Do While .HasNext(Candidate)
                    If Equality.IsTrue(Element, Candidate) Then _
                        Contains = True: Exit Function
                Loop
            
            End With
        
        Case TypeOf This.Sequence Is IGenericCollection
            '@Ignore ArgumentWithIncompatibleObjectType
            Contains = CastCollection(This.Sequence).Contains(Element)
            
        Case Else
            With This.Sequence.Iterator
                If Not (Element Is Nothing) Then
                    
                    Set Casted = Element
                    Do While .HasNext(Candidate)
                        If Casted.Equals(Candidate) Then _
                            Contains = True: Exit Function
                    Loop
                        
                Else
                        
                    Do While .HasNext(Candidate)
                        If (Candidate Is Element) Then _
                            Contains = True: Exit Function
                    Loop
                
                End If
            End With
            
    End Select
    
End Function

Public Function ContainsAll(ByVal Other As IGenericSequence, Optional ByVal Mapper As GenericMapper, Optional ByVal HashProvider As IGenericHashProvider) As Boolean
    
    If (This.Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function ContainsAll").IsNothing("Sequence")
    
    '@Ignore ArgumentWithIncompatibleObjectType
    '@Ignore ArgumentWithIncompatibleObjectType
    If TypeOf This.Sequence Is IGenericCollection Then _
        If TypeOf Other Is IGenericCollection Then _
            If Not CastCollection(This.Sequence).Count = CastCollection(Other).Count Then _
                Exit Function
            
    ContainsAll = Not GenericSequence.Stream(This.Sequence).Except(Other, Mapper, HashProvider).HasAny
    
End Function

Public Function Equals(ByVal Other As IGenericSequence, Optional ByVal Equality As IGenericPredicator) As Boolean

    Dim OtherIterator As IGenericIterator
    Dim x As Object, y As Object
    
    If (This.Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Equals").IsNothing("Sequence")
    
    If (Other Is Nothing) Then _
        Exit Function

    If TypeOf Other Is GenericSequence Then _
        Set Other = GetSource(Other)
    
    If (Other Is This.Sequence) Then _
        Equals = True: Exit Function
    
    If (Equality Is Nothing) Then
    
         Equals = GenericSequence.SequenceEquality.IsTrue(This.Sequence, Other)
         
    Else
    
        Set OtherIterator = Other.Iterator
        
        With This.Sequence.Iterator
            Do While .HasNext(x)
            
                If Not OtherIterator.HasNext(y) Then _
                    Exit Function
            
                If Not Equality.IsTrue(x, y) Then _
                    Exit Function
                
            Loop
        End With
        
        Equals = Not OtherIterator.HasNext(y)
        
    End If
    
End Function

Public Function ToArray() As GenericArray

    Select Case True
        Case (This.Sequence Is Nothing)
            Call GenericError.Build(Me, "Public Function ToArray").IsNothing("Sequence")
        
        Case TypeOf This.Sequence Is IGenericCollection
            '@Ignore ArgumentWithIncompatibleObjectType
            Set ToArray = CastCollection(This.Sequence).ToArray
        
        Case TypeOf This.Sequence Is IGenericSequenceProvider
            Set ToArray = CastProvider(This.Sequence).ToArray
      
        Case Else 'Compute the hard way
            Set ToArray = GenericArray.BuildFrom(This.Sequence)
            
    End Select
    
End Function

Public Function ToList() As GenericOrderedList
    
    Select Case True
        Case (This.Sequence Is Nothing)
            Call GenericError.Build(Me, "Public Function ToList").IsNothing("Sequence")
        
        Case TypeOf This.Sequence Is IGenericSequenceProvider
            Set ToList = CastProvider(This.Sequence).ToList
        
        Case Else 'Compute the hard way
            Set ToList = GenericOrderedList.Build().AddAll(This.Sequence)
            
    End Select
    
End Function

Public Function ToHashSet(Optional ByVal HashProvider As IGenericHashProvider) As GenericHashSet
    
    Dim Capacity As Long

    Select Case True
        Case (This.Sequence Is Nothing)
            Call GenericError.Build(Me, "Public Function ToHashSet").IsNothing("Sequence")
            
        Case TypeOf This.Sequence Is IGenericCollection
            '@Ignore ArgumentWithIncompatibleObjectType
            Capacity = CastCollection(This.Sequence).Count
            
        Case TypeOf This.Sequence Is IGenericSequenceProvider
            Capacity = CastProvider(This.Sequence).GuessCount
            
        Case Else
            Capacity = 128
            
    End Select
    
    Set ToHashSet = GenericHashSet.Build(Capacity:=Capacity, HashProvider:=HashProvider).DoUnion(This.Sequence).TrimExcess
    
End Function

Public Function ToPartitionList(ByVal Size As Long, Optional ByVal Batch As Boolean = False) As GenericOrderedList
    
    Dim i As Long, Count As Long
    Dim Result As GenericOrderedList
    Dim Chunk As GenericArray, tmp As GenericArray
    Dim Element As Object
    
    Select Case True
        Case (This.Sequence Is Nothing)
            Call GenericError.Build(Me, "Public Function ToPartitionList").IsNothing("Sequence")
        
        Case (Size < 0)
            Call GenericError.Build(Me, "Public Function ToPartitionList").ArgOutOfRange("Size < 0")
            
        Case Else
            
            Count = Me.Count 'Cache Count to avoid evaluating it twice
            
            If Count < Size Then
            
                Set Result = GenericOrderedList.Build(1)
                Call Result.Add(Me.Source)
                
            Else
                
                Count = VBA.Fix((Count / Size))
            
                If Not ((Count Mod Size) = 0) Then _
                    Count = Count + 1
                
                Set Result = GenericOrderedList.Build(Count)
                Set ToPartitionList = Result
                
                If Batch Then
                    
                    With Me.Iterator
                        Do While .HasNext(Element)
                            
                            Set Chunk = GenericArray.Build(Size)
                            Set Chunk.ElementAt(0) = Element
                            
                            For i = 1 To Size - 1
                            
                                If Not .HasNext(Element) Then
                                    Set tmp = Chunk
                                    Set Chunk = GenericArray.Build(i)
                                    Call tmp.CopyTo(Target:=Chunk, Index:=Chunk.LowerBound, BeginAt:=tmp.LowerBound, Count:=i)
                                    Call Result.Add(Chunk)
                                    Exit Function
                                End If
                                
                                Set Chunk.ElementAt(i) = Element
                                
                            Next
                            
                            Call Result.Add(Chunk)
                        Loop
                    End With
                   
                    
                Else
                    
                    For i = 0 To Count - 1: Call Result.Add(GenericSequence.Stream(This.Sequence).Skip(i * Size).Take(Size).Source): Next
                
                End If
            End If
    
    End Select
    
End Function

Public Function First() As Object

    Select Case True
        Case (This.Sequence Is Nothing)
            Call GenericError.Build(Me, "Public Function First").IsNothing("Sequence")
        
        Case TypeOf This.Sequence Is IGenericCollection
            '@Ignore ArgumentWithIncompatibleObjectType
            '@Ignore ArgumentWithIncompatibleObjectType
            If CastCollection(This.Sequence).Count > 0 Then _
                Set First = CastCollection(This.Sequence).GetAt(0)
        
        Case TypeOf This.Sequence Is IGenericSequenceProvider
            Set First = CastProvider(This.Sequence).First
            
        Case Else
            Call This.Sequence.Iterator.HasNext(First)

    End Select
    
End Function

Public Function GetAt(ByVal Index As Long) As Object

    Select Case True
        Case (This.Sequence Is Nothing)
            Call GenericError.Build(Me, "Public Function GetAt").IsNothing("Sequence")
        
        Case TypeOf This.Sequence Is IGenericCollection
            '@Ignore ArgumentWithIncompatibleObjectType
            Set GetAt = CastCollection(This.Sequence).GetAt(Index)
           
        Case TypeOf This.Sequence Is IGenericSequenceProvider
            Set GetAt = CastProvider(This.Sequence).GetAt(Index)
                
        Case Else
        
            With This.Sequence.Iterator
                Do While .HasNext(GetAt)
                    If Index = 0 Then _
                        Exit Function
                    Index = Index - 1
                Loop
            End With
            
            Call GenericError.Build(Me, "Public Function GetAt").IndexOutofRange
            
    End Select
        
End Function

Public Function Last() As Object

     Select Case True
        Case (This.Sequence Is Nothing)
            Call GenericError.Build(Me, "Public Function Last").IsNothing("Sequence")
    
        Case TypeOf This.Sequence Is IGenericCollection
            '@Ignore ArgumentWithIncompatibleObjectType
            '@Ignore ArgumentWithIncompatibleObjectType
            '@Ignore ArgumentWithIncompatibleObjectType
            If CastCollection(This.Sequence).Count > 0 Then _
                Set Last = CastCollection(This.Sequence).GetAt(CastCollection(This.Sequence).Count - 1)
        
        Case TypeOf This.Sequence Is IGenericSequenceProvider
            Set Last = CastProvider(This.Sequence).Last
    
        Case Else
            With This.Sequence.Iterator
                '@Ignore EmptyDoWhileBlock
                Do While .HasNext(Last): Loop
            End With
            
    End Select
  
End Function

Public Function Min(Optional ByVal Mapper As GenericMapper) As IGenericValue: Set Min = GenericSequence.Stream(This.Sequence).Descending(Mapper).Last: End Function
Public Function Max(Optional ByVal Mapper As GenericMapper) As IGenericValue: Set Max = GenericSequence.Stream(This.Sequence).Ascending(Mapper).Last: End Function

Public Function IndexOf(ByVal Predicate As IGenericPredicate) As Long
    
    Dim Element As Object
    
    If (This.Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function IndexOf").IsNothing("Sequence")
    
    If (Predicate Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function IndexOf").IsNothing("(Predicate Is Nothing)")
  
    With This.Sequence.Iterator
        Do While .HasNext(Element)
            If Predicate.Match(Element) Then _
                Exit Function
            IndexOf = IndexOf + 1
        Loop
    End With
    
    IndexOf = -1
    
End Function

Public Function Count() As Long
    
    Select Case True
        Case (This.Sequence Is Nothing)
            Call GenericError.Build(Me, "Public Function Count").IsNothing("Sequence")
            
        Case TypeOf This.Sequence Is IGenericCollection
            '@Ignore ArgumentWithIncompatibleObjectType
            Count = CastCollection(This.Sequence).Count
        
        Case TypeOf This.Sequence Is IGenericSequenceProvider
            Count = CastProvider(This.Sequence).Count
       
        Case Else
            With This.Sequence.Iterator
                Do While .HasNext(): Count = Count + 1: Loop
            End With
        
    End Select

End Function

Public Function Sum() As Variant

    Dim Element As Object
    Dim Value As IGenericValue
    Dim Mapper As GenericMapper
    
    Select Case True
        Case (This.Sequence Is Nothing)
            Call GenericError.Build(Me, "Public Function Sum").IsNothing("Sequence")
        
        Case TypeOf This.Sequence Is GenericSequenceMapper
            Set Mapper = SequenceMapper(This.Sequence).Mapper
            
            With SequenceMapper(This.Sequence).Source.Iterator
                If .HasNext(Element) Then
                    
                    Sum = 0
                    
                    Do
                        Set Value = Mapper.Map(Element)
                        If Not Value Is Nothing Then _
                            Sum = Sum + Value.GetValue
                    Loop While .HasNext(Element)
                
                End If
            End With
        
        Case Else
        
            With This.Sequence.Iterator
                If .HasNext(Value) Then
                
                    Sum = 0

                    Do
                        If Not Value Is Nothing Then _
                            Sum = Sum + Value.GetValue
                    Loop While .HasNext(Value)
                    
                End If
            End With
            
    End Select
   
End Function

Public Function Mean() As Variant

    Dim Element As Object
    Dim Value As IGenericValue
    Dim Mapper As GenericMapper
    
    Dim Sum As Variant, Count As Long
    
    Select Case True
        Case (This.Sequence Is Nothing)
            Call GenericError.Build(Me, "Public Function Mean").IsNothing("Sequence")
        
        Case TypeOf This.Sequence Is GenericSequenceMapper
            Set Mapper = SequenceMapper(This.Sequence).Mapper
            
            With SequenceMapper(This.Sequence).Source.Iterator
                If .HasNext(Element) Then
                    
                    Sum = 0
                    
                    Do
                        Set Value = Mapper.Map(Element)
                        If Not Value Is Nothing Then _
                            Count = Count + 1: Sum = Sum + Value.GetValue
                    Loop While .HasNext(Element)
                    
                    Mean = (Sum / Count)
                    
                End If
            End With
        
        Case Else
    
            With This.Sequence.Iterator
                If .HasNext(Value) Then
                
                    Sum = 0
                
                    Do
                        If Not Value Is Nothing Then _
                            Count = Count + 1: Sum = Sum + Value.GetValue
                    Loop While .HasNext(Value)
                    
                    Mean = (Sum / Count)
                
                End If
            End With
        
    End Select
                    
End Function

Public Function Variance() As Variant

    Dim Element As Object
    Dim Value As IGenericValue
    Dim Mapper As GenericMapper
    
    Dim x As Variant, Sum As Variant, Sum2 As Variant, Count As Long
    
    Select Case True
        Case (This.Sequence Is Nothing)
            Call GenericError.Build(Me, "Public Function Variance").IsNothing("Sequence")
    
        Case TypeOf This.Sequence Is GenericSequenceMapper
            Set Mapper = SequenceMapper(This.Sequence).Mapper
            
            With SequenceMapper(This.Sequence).Source.Iterator
                If .HasNext(Element) Then
                    
                    Sum = 0
                    Sum2 = 0
                    
                    Do
                        Set Value = Mapper.Map(Element)
                        If Not Value Is Nothing Then
                            x = Value.GetValue
                            Count = Count + 1
                            Sum = (Sum + x)
                            Sum2 = Sum2 + (x * x)
                        End If
                    Loop While .HasNext(Element)
                    
                    Variance = (Sum2 - ((Sum ^ 2) / Count))
                    
                End If
            End With
       
        Case Else
        
            With This.Sequence.Iterator
                If .HasNext(Value) Then
                    
                    Sum = 0
                    Sum2 = 0
                  
                    Do
                        If Not Value Is Nothing Then
                            x = Value.GetValue
                            Count = Count + 1
                            Sum = (Sum + x)
                            Sum2 = Sum2 + (x * x)
                        End If
                    Loop While .HasNext(Value)
                    
                    Variance = (Sum2 - ((Sum ^ 2) / Count))
                    
                End If
            End With
        
    End Select
        
End Function

Public Function StandardDeviation() As Variant
    
    If (This.Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function StandardDeviation").IsNothing("Sequence")
        
    StandardDeviation = VBA.Sqr(Me.Variance())
    
End Function

Private Function GetSource(ByVal Sequence As GenericSequence) As IGenericSequence: Set GetSource = Sequence.Source: End Function
Private Function CastSequence(ByVal Sequence As IGenericSequence) As IGenericSequence: Set CastSequence = Sequence: End Function
Private Function CastProvider(ByVal Sequence As IGenericSequence) As IGenericSequenceProvider: Set CastProvider = Sequence: End Function
Private Function SequenceSet(ByVal Sequence As GenericSequenceSet) As GenericSequenceSet: Set SequenceSet = Sequence: End Function
Private Function SequenceSkip(ByVal Sequence As GenericSequenceSkip) As GenericSequenceSkip: Set SequenceSkip = Sequence: End Function
Private Function SequenceTake(ByVal Sequence As GenericSequenceTake) As GenericSequenceTake: Set SequenceTake = Sequence: End Function
Private Function SequenceRange(ByVal Sequence As GenericSequenceTake) As GenericSequenceTake: Set SequenceRange = Sequence: End Function
Private Function SequenceDistinct(ByVal Sequence As GenericSequenceDistinct) As GenericSequenceDistinct: Set SequenceDistinct = Sequence: End Function
Private Function SequenceWhere(ByVal Sequence As GenericSequenceWhere) As GenericSequenceWhere: Set SequenceWhere = Sequence: End Function
Private Function SequenceMapper(ByVal Sequence As GenericSequenceMapper) As GenericSequenceMapper: Set SequenceMapper = Sequence: End Function
Private Function SequenceSorter(ByVal Sequence As GenericSequenceSorter) As GenericSequenceSorter: Set SequenceSorter = Sequence: End Function
Private Function CastCollection(ByVal Collection As IGenericCollection) As IGenericCollection: Set CastCollection = Collection: End Function

'IGenericHashProvider
Private Function IGenericHashProvider_HashOf(ByVal obj As Object) As Long
    
    Const MaxValue As Long = 2 ^ 31 - 1
    Dim h As Double, HashCode As Long
    Dim Element As Object
    Dim Casted As IGenericValue
    
    If (obj Is Nothing) Then _
        Exit Function
    
    With CastSequence(obj).Iterator
        Do While .HasNext(Element)
            If Not Element Is Nothing Then
                If TypeOf Element Is IGeneric Then
                    Set Casted = Element
                    HashCode = Casted.HashCode
                Else
                    HashCode = ObjPtr(Element)
                End If
                h = 31 * h + HashCode
                IGenericHashProvider_HashOf = CLng(h - (VBA.Fix(h / MaxValue) * MaxValue))
            End If
        Loop
    End With
    
End Function

'Does deep equal
Private Function IGenericHashProvider_Equality(ByVal A As Object, ByVal B As Object) As Boolean: IGenericHashProvider_Equality = GenericSequence.SequenceEquality.IsTrue(A, B): End Function
'IGenericPredicator
Private Function IGenericPredicator_IsTrue(ByVal A As Object, ByVal B As Object) As Boolean
    
    Dim IteratorA As IGenericIterator, IteratorB As IGenericIterator
    Dim x As Object, y As Object
    Dim Casted As IGenericValue
    
    If A Is B Then _
        IGenericPredicator_IsTrue = True: Exit Function
        
    If A Is Nothing Then _
        Exit Function
        
    If B Is Nothing Then _
        Exit Function
        
    If TypeOf A Is IGenericCollection Then _
        If TypeOf B Is IGenericCollection Then _
            If Not (CastCollection(A).Count = CastCollection(B).Count) Then _
                Exit Function 'Early out
                
    'Try cast
    Set IteratorA = CastSequence(A).Iterator
    Set IteratorB = CastSequence(B).Iterator
    
    With IteratorA
        Do While .HasNext(x)
        
            If Not IteratorB.HasNext(y) Then _
                Exit Function
        
            If Not (x Is y) Then
            
                If (x Is Nothing) Then _
                    Exit Function
                
                Set Casted = x
                If Not Casted.Equals(y) Then _
                    Exit Function
                
            End If
            
        Loop
    End With
    
    IGenericPredicator_IsTrue = Not IteratorB.HasNext(y)
    
End Function


