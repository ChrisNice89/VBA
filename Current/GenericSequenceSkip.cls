VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSequenceSkip"
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

Private Const CACHE_START_SIZE As Long = 2 ^ 12

Private Type Member
    Source As IGenericSequence
    Skip As Long
End Type
Private This As Member

Public Property Get Source() As IGenericSequence: Set Source = This.Source: End Property
Public Property Get Skip() As Long: Skip = This.Skip: End Property

Private Function IGenericSequence_Iterator() As IGenericIterator
   
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequence_Iterator").IsNothing("Sequence")
    
    Set IGenericSequence_Iterator = Iterate(This.Source, This.Skip)

End Function

'@Ignore ProcedureNotUsed
Public Function Of(ByVal Source As IGenericSequence, Optional ByVal Skip As Long = 0) As IGenericSequence
    
    Dim Other As GenericSequenceSkip
    
    If (Source Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Of").IsNothing("Sequence")
    
    If Not Me Is GenericSequenceSkip Then _
        Call GenericError.Build(Me, "Public Function Of").InvalidOperation("instance call")
        
    If TypeOf Source Is GenericSequenceSkip Then
        Set Other = Source
        Set Source = Other.Source
        Skip = Skip + Other.Skip
    End If
    
    Set This.Source = Source
    This.Skip = Skip

    Set Of = New GenericSequenceSkip
    This = EmptyStructure
    
End Function
'@Ignore NonReturningFunction, EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Sub Class_Initialize()

    Set This.Source = GenericSequenceSkip.Source
    This.Skip = GenericSequenceSkip.Skip
    
End Sub

Private Function IGenericSequenceProvider_Count() As Long
    
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    Dim Count As Long, i As Long
    
    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "IGenericSequenceProvider_Count").IsNothing("Sequence")
            
        Case (TypeOf This.Source Is IGenericCollection)
            '@Ignore SetAssignmentWithIncompatibleObjectType
            Set Collection = This.Source
            Count = Collection.Count - This.Skip
            If (Count < 0) Then _
                Count = 0
        
        Case (TypeOf This.Source Is IGenericSequenceProvider) 'its a full iteration in any case
            Set Provider = This.Source
            Count = Provider.Count - This.Skip
            If (Count < 0) Then _
                Count = 0
                
        Case Else
                
                i = This.Skip - 1 'reduce 1 for a tight loop
                
                With This.Source.Iterator
                    If (This.Skip > 0) Then _
                        Do While .HasNext And i > 0: i = i - 1: Loop
                    Do While .HasNext: Count = Count + 1: Loop
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
            guess = Collection.Count - This.Skip
            If (guess < 0) Then _
                guess = 0
                
        Case (TypeOf This.Source Is IGenericSequenceProvider)
            Set Provider = This.Source
            guess = Provider.GuessCount - This.Skip
            If (guess < 0) Then _
                guess = 0
             
        Case Else
            guess = -1
            
    End Select
    
    IGenericSequenceProvider_GuessCount = guess
    
End Function

Private Function IGenericSequenceProvider_First() As Object
    
    Dim Collection As IGenericCollection
    Dim Remaining As Long
   
    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_First").IsNothing("This.Source")
        
        Case (TypeOf This.Source Is IGenericCollection)
            '@Ignore SetAssignmentWithIncompatibleObjectType
            Set Collection = This.Source
            If (Collection.Count > This.Skip) Then _
                Set IGenericSequenceProvider_First = Collection.GetAt(This.Skip)
           
        Case Else
            
            Remaining = This.Skip
            
            With This.Source.Iterator
                Do While (Remaining > 0)
                    If Not .HasNext Then _
                        Exit Function
                    Remaining = Remaining - 1
                Loop
                
                Call .HasNext(IGenericSequenceProvider_First)
                
            End With
            
    End Select
   
End Function

Private Function IGenericSequenceProvider_Last() As Object
    
    Dim Collection As IGenericCollection, Remaining As Long

    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_Last").IsNothing("Sequence")
    
        Case TypeOf This.Source Is IGenericCollection
            '@Ignore SetAssignmentWithIncompatibleObjectType
            Set Collection = This.Source
            If (Collection.Count > This.Skip) Then _
                Set IGenericSequenceProvider_Last = Collection.GetAt(Collection.Count - 1)
        'cant go for a provider, because skip is maybe > sequence.count
        Case Else
            
            Remaining = This.Skip
            
            With This.Source.Iterator
                Do While (Remaining > 0)
                    If Not .HasNext Then _
                        Exit Function
                    Remaining = Remaining - 1
                Loop
                
                '@Ignore EmptyDoWhileBlock
                Do While .HasNext(IGenericSequenceProvider_Last): Loop
        
            End With
    
    End Select
    
End Function

Private Function IGenericSequenceProvider_GetAt(ByVal Index As Long) As Object
    
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    
    Dim Element As Object
    Dim Remaining As Long
    
    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GetAt").IsNothing("Sequence")
       
        Case TypeOf This.Source Is IGenericCollection
            '@Ignore SetAssignmentWithIncompatibleObjectType
            Set Collection = This.Source
            Set IGenericSequenceProvider_GetAt = Collection.GetAt(This.Skip + Index)
        
        Case TypeOf This.Source Is IGenericSequenceProvider
            Set Provider = This.Source
            Set IGenericSequenceProvider_GetAt = Provider.GetAt(This.Skip + Index)
           
        Case Else
            
            Remaining = This.Skip
            
            With This.Source.Iterator
                Do While (Remaining > 0)
                    If Not .HasNext Then _
                        Exit Function
                    Remaining = Remaining - 1
                Loop
                
                Do While .HasNext(Element)
                    If (Index = 0) Then
                        Set IGenericSequenceProvider_GetAt = Element
                        Exit Function
                    End If
                    Index = Index - 1
                Loop
            End With
            
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GetAt").IndexOutofRange
    
    End Select
    
End Function

Private Function IGenericSequenceProvider_ToArray() As GenericArray
    
    Dim Collection As IGenericCollection
    Dim Provider As IGenericSequenceProvider
    
    Dim Element As Object
    Dim Result As GenericArray, tmp As GenericArray
    
    Dim Capacity As Long, Count As Long
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_ToArray").IsNothing("Sequence")
    
    Set Result = GenericArray ' Avoid null
    
    If (This.Skip > 0) Then
        If (TypeOf This.Source Is IGenericCollection) Then 'fast Path
            
            '@Ignore SetAssignmentWithIncompatibleObjectType
            Set Collection = This.Source
            Capacity = Collection.Count
            
            If (Capacity > This.Skip) Then
                
                Set Result = GenericArray.Build(Capacity - This.Skip)
                
                If (TypeOf This.Source Is GenericArray) Then 'we can do even better
                    
                    Set tmp = This.Source
                    Call tmp.CopyTo(Target:=Result, Index:=Result.LowerBound, BeginAt:=This.Skip, Count:=Result.Length)
                    
                Else
                    
                    With Iterate(This.Source, This.Skip)
                        Do While .HasNext(Element): Set Result.ElementAt(Count) = Element: Count = Count + 1: Loop
                    End With
                    
                End If
            End If
        
        Else
            
            With Iterate(This.Source, This.Skip)
                If .HasNext(Element) Then
                            
                    If (TypeOf This.Source Is IGenericSequenceProvider) Then
                        Set Provider = This.Source
                        Capacity = Provider.GuessCount - This.Skip
                    End If
            
                    If (Capacity <= 0) Then _
                        Capacity = CACHE_START_SIZE
                        
                    Set Result = GenericArray.Build(Capacity)
                    
                    Do
                    
                        If (Count = Capacity) Then
                            Set tmp = Result
                            Set Result = GenericArray.Build(Capacity * 2)
                            Call tmp.CopyTo(Target:=Result, Index:=Result.LowerBound, BeginAt:=tmp.LowerBound, Count:=Count)
                            Capacity = Result.Length
                        End If
                        
                        Set Result.ElementAt(Count) = Element
                        Count = Count + 1
                        
                    Loop While .HasNext(Element)
        
                    If (Count < Result.Length) Then
                        Set tmp = Result
                        Set Result = GenericArray.Build(Count)
                        Call tmp.CopyTo(Target:=Result, Index:=Result.LowerBound, BeginAt:=tmp.LowerBound, Count:=Count)
                    End If
                
                End If
            End With
            
        End If
    End If
    
    Set IGenericSequenceProvider_ToArray = Result
    
End Function

Private Function IGenericSequenceProvider_ToList() As GenericOrderedList
     
    Dim Element As Object
    Dim Result As GenericOrderedList
    Dim Capacity As Long

    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_ToList").IsNothing("Sequence")
    
    Set Result = GenericOrderedList.Build
    
    If (This.Skip > 0) Then
        With Iterate(This.Source, This.Skip)
            If .HasNext(Element) Then
                
                Capacity = IGenericSequenceProvider_GuessCount
                If (Capacity < 0) Then _
                    Capacity = CACHE_START_SIZE
                
                Call Result.EnsureCapacity(Capacity)
                
                Do: Call Result.Add(Element): Loop While .HasNext(Element)
                
            End If
        End With
    End If
    
    Set IGenericSequenceProvider_ToList = Result
    
End Function

Private Function Iterate(ByVal Source As IGenericSequence, ByVal Index As Long) As IGenericIterator
    
    Dim Collection As IGenericCollection
    
    Set Iterate = Me 'Avoid null
    
    If (Index > 0) Then
        If TypeOf Source Is IGenericCollection Then
            '@Ignore SetAssignmentWithIncompatibleObjectType
            Set Collection = Source
            If (Index >= Collection.Count) Then _
                Exit Function
             
            If Collection.HasRandomAccess Then _
                Set Iterate = GenericIndexIterator.Build(Collection, Index)
        
        Else
                
            Set Iterate = Source.Iterator
            With Iterate
                Do While .HasNext
                    Index = Index - 1
                    If (Index = 0) Then _
                        Exit Do
                Loop
            End With
        
        End If
    End If
    
End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As Object: Set IGenericIterator_Current = Nothing: End Property
Private Function IGenericIterator_HasNext(Optional Out As Object) As Boolean: IGenericIterator_HasNext = False: End Function


