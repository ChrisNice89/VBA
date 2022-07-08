VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSequenceWhere"
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
    Predicate As IGenericPredicate
End Type
Private This As Member

Public Property Get Source() As IGenericSequence: Set Source = This.Source: End Property
Public Property Get Predicate() As IGenericPredicate: Set Predicate = This.Predicate: End Property

Private Function IGenericSequence_Iterator() As IGenericIterator
   
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequence_Iterator").IsNothing("Sequence")
    
    If (This.Predicate Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequence_Iterator").IsNothing("Predicate")
    
    Set This.Iterator = This.Source.Iterator
        
    Set IGenericSequence_Iterator = Me
    
End Function

Public Function Of(ByVal Source As IGenericIterator, ByVal Predicate As IGenericPredicate) As IGenericSequence

    If (Source Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Of").IsNothing("Sequence")
        
    If (Predicate Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Of").IsNothing("Predicate")
    
    If Not Me Is GenericSequenceWhere Then _
        Call GenericError.Build(Me, "Public Function Of").InvalidOperation("instance call")
  
    Set This.Source = Source
    Set This.Predicate = Predicate
    
    Set Of = New GenericSequenceWhere
    This = EmptyStructure
    
End Function
'@Ignore NonReturningFunction, EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Sub Class_Initialize()

    Set This.Source = GenericSequenceWhere.Source
    Set This.Predicate = GenericSequenceWhere.Predicate
        
End Sub

Private Function IGenericSequenceProvider_Count() As Long
    
    '@Ignore VariableNotAssigned
    Dim Element As Object
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_Count").IsNothing("Sequence")
     
    With This.Source.Iterator
        Do While .HasNext(Element)
            If This.Predicate.Match(Element) Then _
                IGenericSequenceProvider_Count = IGenericSequenceProvider_Count + 1
        Loop
    End With
 
End Function

Private Function IGenericSequenceProvider_GuessCount() As Long
    
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    Dim guess As Long
    
    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GuessCount").IsNothing("Sequence")
            
        Case (TypeOf This.Source Is IGenericCollection)
            Set Collection = This.Source
            guess = Collection.Count
        
        Case (TypeOf This.Source Is IGenericSequenceProvider)
            Set Provider = This.Source
            guess = Provider.GuessCount
           
        Case Else
            guess = -1
            
    End Select
    
    IGenericSequenceProvider_GuessCount = guess

End Function

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
   
    '@Ignore VariableNotAssigned
    Dim Element As Object
   
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_Last").IsNothing("Sequence")
    
    With This.Source.Iterator
        Do While .HasNext(Element)
            If This.Predicate.Match(Element) Then _
                Set IGenericSequenceProvider_Last = Element
        Loop
    End With
            
End Function

Private Function IGenericSequenceProvider_GetAt(ByVal Index As Long) As Object

    '@Ignore VariableNotAssigned
    Dim Element As Object
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GetAt").IsNothing("Sequence")
            
    With This.Source.Iterator
        Do While .HasNext(Element)
            If This.Predicate.Match(Element) Then
                If (Index = 0) Then
                    Set IGenericSequenceProvider_GetAt = Element
                    Exit Function
                End If
                Index = Index - 1
            End If
        Loop
    End With
    
    Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GetAt").IndexOutofRange
    
End Function

Private Function IGenericSequenceProvider_ToArray() As GenericArray
  
    '@Ignore VariableNotAssigned
    Dim Element As Object
    Dim Result As GenericArray, tmp As GenericArray
    
    Dim Capacity As Long, Count As Long
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_ToArray").IsNothing("Sequence")
    
    Set Result = GenericArray ' Avoid null
    
    With This.Source.Iterator
        Do While .HasNext(Element)
            If This.Predicate.Match(Element) Then
            
                If (Count = Capacity) Then
                    If (Capacity > 0) Then

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
        Loop
        
        If (Count < Result.Length) Then
            Set tmp = Result
            Set Result = GenericArray.Build(Count)
            Call tmp.CopyTo(Target:=Result, Index:=Result.LowerBound, BeginAt:=tmp.LowerBound, Count:=Count)
        End If
    
    End With
    
    Set IGenericSequenceProvider_ToArray = Result
    
End Function

Private Function IGenericSequenceProvider_ToList() As GenericOrderedList
     
    '@Ignore VariableNotAssigned
    Dim Element As Object
    Dim Result As GenericOrderedList
    Dim Capacity As Long
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_ToList").IsNothing("Sequence")
    
    Set Result = GenericOrderedList.Build
    
    With This.Source.Iterator
        Do While .HasNext(Element)
            If This.Predicate.Match(Element) Then
            
                If (Capacity = 0) Then
                
                    Capacity = IGenericSequenceProvider_GuessCount
                    If (Capacity <= 0) Then _
                        Capacity = CACHE_START_SIZE
                        
                    Call Result.EnsureCapacity(Capacity)
                        
                End If
                
                Call Result.Add(Element)
                
            End If
        Loop
    End With
    
    Set IGenericSequenceProvider_ToList = Result.TrimExcess
    
End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As Object: Set IGenericIterator_Current = This.Iterator.Current: End Property

Private Function IGenericIterator_HasNext(Optional Out As Object) As Boolean
   
    '@Ignore VariableNotAssigned
    Dim Element As Object
    
    With This.Iterator
        Do While .HasNext(Element)
            If This.Predicate.Match(Element) Then
                Set Out = Element
                IGenericIterator_HasNext = True
                Exit Function
            End If
        Loop
    End With
    
    Set This.Iterator = Nothing
    
End Function


