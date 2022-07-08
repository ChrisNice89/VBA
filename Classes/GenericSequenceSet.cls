VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSequenceSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IGenericSequenceProvider
Implements IGenericSequence
Implements IGenericIterator

Private Const CACHE_START_SIZE As Long = 2 ^ 12

Private Type Member
    Source As IGenericSequence
    Iterator As IGenericIterator
    HashProvider As IGenericHashProvider
    Cache As IGenericCollection
    Mapper As GenericMapper
    Other As IGenericSequence
    Complement As Boolean
End Type
Private This As Member

Public Property Get Source() As IGenericSequence: Set Source = This.Source: End Property
Public Property Get IsComplement() As Boolean: IsComplement = This.Complement: End Property
Public Property Get Other() As IGenericSequence: Set Other = This.Other: End Property
Public Property Get HashProvider() As IGenericHashProvider: Set HashProvider = This.HashProvider: End Property
Public Property Get Mapper() As GenericMapper: Set Mapper = This.Mapper: End Property

Private Function IGenericSequence_Iterator() As IGenericIterator
   
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequence_Iterator").IsNothing("Sequence")
    
    Set This.Iterator = This.Source.Iterator
    Set This.Cache = CreateCache(This.Other, This.HashProvider)
    
    If This.Cache.Count > 0 Then
        Set IGenericSequence_Iterator = Me
    Else
        Set IGenericSequence_Iterator = This.Iterator
    End If
    
End Function

Public Function Of( _
                        ByVal Source As IGenericSequence, _
                        ByVal Other As IGenericSequence, _
                        Optional ByVal Mapper As GenericMapper, _
                        Optional ByVal HashProvider As IGenericHashProvider, _
                        Optional ByVal Complement As Boolean) As IGenericSequence

    If (Source Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Of").IsNothing("Sequence")
        
    If (Other Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Of").IsNothing("Args")
    
    If Not Me Is GenericSequenceSet Then _
        Call GenericError.Build(Me, "Public Function Of").InvalidOperation("instance call")
    
    If (Mapper Is Nothing) Then _
        Set Mapper = GenericMapper.Build
        
    This.Complement = Complement
    Set This.Source = Source
    Set This.Other = Other
    Set This.Mapper = Mapper
    Set This.HashProvider = HashProvider
    
    Set Of = New GenericSequenceSet
    This = EmptyStructure
    
End Function
'@Ignore NonReturningFunction, EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Sub Class_Initialize()
    
    With GenericSequenceSet
        This.Complement = .IsComplement
        Set This.Source = .Source
        Set This.Other = .Other
        Set This.Iterator = IGenericIterator ' Avoid Null
        Set This.HashProvider = .HashProvider
        Set This.Mapper = .Mapper
    End With
    
End Sub

Private Function CreateCache(ByVal Sequence As IGenericSequence, ByVal HashProvider As IGenericHashProvider) As IGenericCollection

    Dim Capacity As Long
    Dim HashSet As GenericHashSet
    
    Set HashSet = GenericHashSet.Build(HashProvider:=HashProvider)
    
    Capacity = GuessCount(Sequence)
    If (Capacity < 0) Then _
        Capacity = CACHE_START_SIZE
    
    Call HashSet.EnsureCapacity(Capacity).DoUnion(Sequence)
    
    Set CreateCache = HashSet
    
End Function

Private Function GuessCount(ByVal Sequence As IGenericSequence) As Long
    
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    
    Select Case True
        Case (Sequence Is Nothing)
            Call GenericError.Build(Me, "Private Function GuessCount").IsNothing("Sequence")
            
        Case (TypeOf Sequence Is IGenericCollection)
            Set Collection = Sequence
            GuessCount = Collection.Count
        
        Case (TypeOf Sequence Is IGenericSequenceProvider)
            Set Provider = Sequence
            GuessCount = Provider.GuessCount
           
        Case Else
            GuessCount = -1
            
    End Select
    
End Function

Private Function IGenericSequenceProvider_Count() As Long
    
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    
    '@Ignore VariableNotAssigned
    Dim Element As Object
    Dim Cache As IGenericCollection
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_Count").IsNothing("Sequence")
        
    Set Cache = CreateCache(This.Other, This.HashProvider)
    
    If This.Cache.Count > 0 Then
        
        With This.Source.Iterator
            Do While .HasNext(Element)
                If Not (Cache.Contains(This.Mapper.Map(Element)) = This.Complement) Then _
                    IGenericSequenceProvider_Count = IGenericSequenceProvider_Count + 1
            Loop
        End With
        
    Else
    
        Select Case True
                
            Case (TypeOf This.Source Is IGenericCollection)
                Set Collection = This.Source
                IGenericSequenceProvider_Count = Collection.Count
            
            Case (TypeOf This.Source Is IGenericSequenceProvider)
                Set Provider = This.Source
                IGenericSequenceProvider_Count = Provider.Count
               
            Case Else
                With This.Source.Iterator
                    Do While .HasNext(): IGenericSequenceProvider_Count = IGenericSequenceProvider_Count + 1: Loop
                End With
                
        End Select
    
    End If
    
End Function

Private Function IGenericSequenceProvider_GuessCount() As Long: IGenericSequenceProvider_GuessCount = GuessCount(This.Source): End Function

'@Ignore NonReturningFunction
Private Function IGenericSequenceProvider_First() As Object
    
    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_First").IsNothing("This.Source")
           
        Case Else
            Call IGenericSequence_Iterator.HasNext(IGenericSequenceProvider_First)
            
    End Select
  
End Function

Private Function IGenericSequenceProvider_Last() As Object
    
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    
    '@Ignore VariableNotAssigned
    Dim Element As Object
    Dim Cache As IGenericCollection
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_Last").IsNothing("Sequence")

    Set Cache = CreateCache(This.Other, This.HashProvider)
    
    If This.Cache.Count > 0 Then
    
        With This.Source.Iterator
            Do While .HasNext(Element)
                If Not (Cache.Contains(This.Mapper.Map(Element)) = This.Complement) Then _
                    Set IGenericSequenceProvider_Last = Element
            Loop
        End With
        
    Else
        
        Select Case True
                
            Case (TypeOf This.Source Is IGenericCollection)
                Set Collection = This.Source
                If (Collection.Count > 0) Then _
                    Set IGenericSequenceProvider_Last = Collection.GetAt(Collection.Count - 1)
            
            Case (TypeOf This.Source Is IGenericSequenceProvider)
                Set Provider = This.Source
                Set IGenericSequenceProvider_Last = Provider.Last
               
            Case Else
                With This.Source.Iterator
                    '@Ignore EmptyDoWhileBlock
                    Do While .HasNext(IGenericSequenceProvider_Last):  Loop
                End With
                
        End Select

    End If
    
End Function

Private Function IGenericSequenceProvider_GetAt(ByVal Index As Long) As Object

    '@Ignore VariableNotAssigned
    Dim Element As Object
    Dim Cache As IGenericCollection
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GetAt").IsNothing("Sequence")
            
    With This.Source.Iterator
        If .HasNext(Element) Then
            
            Set Cache = CreateCache(This.Other, This.HashProvider)
            
            If This.Cache.Count > 0 Then
                Do
                    If Not (Cache.Contains(This.Mapper.Map(Element)) = This.Complement) Then
                        If (Index = 0) Then
                            Set IGenericSequenceProvider_GetAt = Element
                            Exit Function
                        End If
                        Index = Index - 1
                    End If
                Loop While .HasNext(Element)
            End If
            
        End If
    End With
    
    Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GetAt").IndexOutofRange
    
End Function

Private Function IGenericSequenceProvider_ToArray() As GenericArray
    
    '@Ignore VariableNotAssigned
    Dim Element As Object
    Dim Result As GenericArray, tmp As GenericArray
    Dim Cache As IGenericCollection
    
    Dim Capacity As Long, Count As Long
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_ToArray").IsNothing("Sequence")
    
    Set Result = GenericArray ' Avoid null
    
    With This.Source.Iterator
        If .HasNext(Element) Then
                    
            Set Cache = CreateCache(This.Other, This.HashProvider)
            
            If Cache.Count > 0 Then
                Do
                    If Not (Cache.Contains(This.Mapper.Map(Element)) = This.Complement) Then
                    
                        If (Count = Capacity) Then
                            If Capacity > 0 Then

                                Set tmp = Result
                                Set Result = GenericArray.Build(Capacity * 2)
                                Call tmp.CopyTo(Target:=Result, Index:=Result.LowerBound, BeginAt:=tmp.LowerBound, Count:=Count)
                                Capacity = Result.Length
                                
                            Else
                            
                                Capacity = IGenericSequenceProvider_GuessCount
                                If (Capacity <= 0) Then _
                                    Capacity = CACHE_START_SIZE
                                Set Result = GenericArray.Build(Capacity)
                                
                            End If
                        End If
                        
                        Set Result.ElementAt(Count) = Element
                        Count = Count + 1
                  
                    End If
                Loop While .HasNext(Element)
            
            Else
                
                Do
                 
                    If (Count = Capacity) Then
                        If Capacity > 0 Then

                            Set tmp = Result
                            Set Result = GenericArray.Build(Capacity * 2)
                            Call tmp.CopyTo(Target:=Result, Index:=Result.LowerBound, BeginAt:=tmp.LowerBound, Count:=Count)
                            Capacity = Result.Length
                            
                        Else
                        
                            Capacity = IGenericSequenceProvider_GuessCount
                            If (Capacity <= 0) Then _
                                Capacity = CACHE_START_SIZE
                            Set Result = GenericArray.Build(Capacity)
                            
                        End If
                    End If
                    
                    Set Result.ElementAt(Count) = Element
                    Count = Count + 1
                  
                Loop While .HasNext(Element)
            
            End If
            
            If (Count < Result.Length) Then
                Set tmp = Result
                Set Result = GenericArray.Build(Count)
                Call tmp.CopyTo(Target:=Result, Index:=Result.LowerBound, BeginAt:=tmp.LowerBound, Count:=Count)
            End If
    
        End If
    End With
    
    Set IGenericSequenceProvider_ToArray = Result
    
End Function

Private Function IGenericSequenceProvider_ToList() As GenericOrderedList
    
    '@Ignore VariableNotAssigned
    Dim Element As Object
    Dim Result As GenericOrderedList
    Dim Capacity As Long
    Dim Cache As IGenericCollection
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_ToList").IsNothing("Sequence")
    
    Set Result = GenericOrderedList.Build
    
    With This.Source.Iterator
        If .HasNext(Element) Then
            
            Capacity = IGenericSequenceProvider_GuessCount
            If (Capacity < 0) Then _
                Capacity = CACHE_START_SIZE
                
            Call Result.EnsureCapacity(Capacity)
                
            Set Cache = CreateCache(This.Other, This.HashProvider)
            
            If (Cache.Count > 0) Then
            
                Do
                    If Not (Cache.Contains(This.Mapper.Map(Element)) = This.Complement) Then _
                        Call Result.Add(Element)
                Loop While .HasNext(Element)
                
            Else
            
                Do: Call Result.Add(Element): Loop While .HasNext(Element)
                
            End If
            
        End If
    End With
    
    Set IGenericSequenceProvider_ToList = Result.TrimExcess
    
End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As Object: Set IGenericIterator_Current = This.Iterator.Current: End Property

Private Function IGenericIterator_HasNext(Optional Out As Object) As Boolean
    
    Do While This.Iterator.HasNext(Out)
        If Not (This.Cache.Contains(This.Mapper.Map(Out)) = This.Complement) Then _
            IGenericIterator_HasNext = True: Exit Function
    Loop

    Set This.Cache = Nothing
    Set This.Iterator = Nothing
    
End Function


