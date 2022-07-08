VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSequenceTake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IGenericSequenceProvider
Implements IGenericSequence
Implements IGenericIterator

Private Type Member
    Source As IGenericSequence
    Iterator As IGenericIterator
    Take As Long
    Remaining As Long
End Type
Private This As Member

Public Property Get Source() As IGenericSequence: Set Source = This.Source: End Property
Public Property Get Take() As Long: Take = This.Take: End Property

Private Function IGenericSequence_Iterator() As IGenericIterator
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequence_Iterator").IsNothing("Sequence")
    
    Set IGenericSequence_Iterator = Me
    
    If (This.Take > 0) Then _
        Set This.Iterator = This.Source.Iterator
    
    This.Remaining = This.Take
    Set IGenericSequence_Iterator = Me
       
End Function

Public Function Of(ByVal Source As IGenericSequence, ByVal Take As Long) As IGenericSequence
    
    Dim Other As GenericSequenceTake
    
    If (Source Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Of").IsNothing("Sequence")
    
    If Not Me Is GenericSequenceTake Then _
        Call GenericError.Build(Me, "Public Function Of").InvalidOperation("instance call")
        
    If TypeOf Source Is GenericSequenceTake Then
    
        Set Other = Source
        Set Source = Other.Source
     
        If Other.Take < Take Then _
            Take = Other.Take
           
    End If
    
    Set This.Source = Source
    This.Take = Take
    
    Set Of = New GenericSequenceTake
    This = EmptyStructure
    
End Function
'@Ignore NonReturningFunction, EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Sub Class_Initialize()
    
    Set This.Source = GenericSequenceTake.Source
    This.Take = GenericSequenceTake.Take
   
End Sub

Private Function IGenericSequenceProvider_Count() As Long
    Dim Collection As IGenericCollection
    Dim Count As Long, Remaining As Long

    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "IGenericSequenceProvider_Count").IsNothing("Sequence")
            
        Case (TypeOf This.Source Is IGenericCollection)
            Set Collection = This.Source
            Count = Collection.Count
            If (Count > This.Take) Then _
                Count = This.Take
           
        Case Else
                 
            Remaining = This.Take
    
            With This.Source.Iterator
                Do While (Remaining > 0)
                    If Not .HasNext() Then _
                        Exit Do
                    Count = Count + 1
                    Remaining = Remaining - 1
                Loop
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
            Set Collection = This.Source
            guess = Collection.Count
        
        Case (TypeOf This.Source Is IGenericSequenceProvider)
            Set Provider = This.Source
            guess = Provider.GuessCount
           
        Case Else
            guess = -1
            
    End Select
    
    IGenericSequenceProvider_GuessCount = VBA.IIf(guess > This.Take, This.Take, guess)
    
End Function

Private Function IGenericSequenceProvider_First() As Object
    
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
   
    If (This.Take > 0) Then
        Select Case True
            Case (This.Source Is Nothing)
                Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_First").IsNothing("This.Source")
            
            Case (TypeOf This.Source Is IGenericCollection)
                Set Collection = This.Source
                If (Collection.Count > 0) Then _
                    Set IGenericSequenceProvider_First = Collection.GetAt(0)
            
            Case TypeOf This.Source Is IGenericSequenceProvider
                Set Provider = This.Source
                Set IGenericSequenceProvider_First = Provider.First
            
            Case Else
                Call This.Source.Iterator.HasNext(IGenericSequenceProvider_First)
                
        End Select
    End If
    
End Function

Private Function IGenericSequenceProvider_Last() As Object
    Dim Collection As IGenericCollection, Remaining As Long
    
    If (This.Take > 0) Then
        Select Case True
            Case (This.Source Is Nothing)
                Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_Last").IsNothing("Sequence")
        
            Case TypeOf This.Source Is IGenericCollection
                Set Collection = This.Source
                
                If (Collection.Count > This.Take) Then
                    Set IGenericSequenceProvider_Last = Collection.GetAt(This.Take - 1)
                Else
                    Set IGenericSequenceProvider_Last = Collection.GetAt(Collection.Count - 1)
                End If
            'cant go for a provider, because take is maybe > sequence.count
            Case Else
                
                Remaining = This.Take
                
                With This.Source.Iterator
                    Do While (Remaining > 0)
                        If Not .HasNext(IGenericSequenceProvider_Last) Then _
                            Exit Function
                        Remaining = Remaining - 1
                    Loop
                End With
               
            End Select
        End If
        
End Function

Private Function IGenericSequenceProvider_GetAt(ByVal Index As Long) As Object
    
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    '@Ignore VariableNotAssigned
    Dim Element As Object
    
    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GetAt").IsNothing("Sequence")
            
        Case (Index >= This.Take)
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GetAt").IndexOutofRange
            
        Case TypeOf This.Source Is IGenericCollection
            Set Collection = This.Source
            Set IGenericSequenceProvider_GetAt = Collection.GetAt(Index)
        
        Case TypeOf This.Source Is IGenericSequenceProvider
            Set Provider = This.Source
            Set IGenericSequenceProvider_GetAt = Provider.GetAt(Index)
           
        Case Else
            With This.Source.Iterator
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
    Dim Result As GenericArray, tmp As GenericArray
    '@Ignore VariableNotAssigned
    Dim Element As Object
    
    Dim Capacity As Long, Count As Long, Remaining As Long
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_ToArray").IsNothing("Sequence")
    
    Set Result = GenericArray 'Avoid null
    Remaining = This.Take
    
    If (TypeOf This.Source Is IGenericCollection) Then
    
        Set Collection = This.Source
        Capacity = Collection.Count
        
        If (Capacity <= Remaining) Then 'Fast path
            Set Result = GenericArray.Build(Capacity)
            Call Collection.CopyTo(Target:=Result, Index:=Result.LowerBound) 'Fetch all
            Remaining = 0
        End If
        
    End If
        
    If (Remaining > 0) Then
        With This.Source.Iterator
            If .HasNext(Element) Then
            
                Set Result = GenericArray.Build(Remaining)
                
                Do
                    
                    Set Result.ElementAt(Count) = Element
                    Count = Count + 1
                    
                    Remaining = Remaining - 1
                    If (Remaining = 0) Then _
                        Exit Do
                    
                Loop While .HasNext(Element)
      
                If (Count < Result.Length) Then
                    Set tmp = Result
                    Set Result = GenericArray.Build(Count)
                    Call tmp.CopyTo(Target:=Result, Index:=Result.LowerBound, BeginAt:=tmp.LowerBound, Count:=Count)
                End If
            
            End If
        End With
    End If
    
    Set IGenericSequenceProvider_ToArray = Result
    
End Function

Private Function IGenericSequenceProvider_ToList() As GenericOrderedList
     
    Dim Collection As IGenericCollection
    Dim Result As GenericOrderedList
    '@Ignore VariableNotAssigned
    Dim Element As Object
    Dim Remaining As Long
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_ToList").IsNothing("Sequence")
    
    Set Result = GenericOrderedList.Build
    Remaining = This.Take
    
    If (TypeOf This.Source Is IGenericCollection) Then
        Set Collection = This.Source
        
        If (Collection.Count <= Remaining) Then 'Fast path
            Call Result.AddAll(Collection) 'Fetch all
            Remaining = 0
        End If
        
    End If
    
    If (Remaining > 0) Then '!
        With This.Source.Iterator
            If .HasNext(Element) Then
                        
                Call Result.EnsureCapacity(Remaining)
            
                Do
                    
                    Call Result.Add(Element)
                    Remaining = Remaining - 1
                    
                    If (Remaining = 0) Then _
                        Exit Do
                        
                Loop While .HasNext(Element)
            
            End If
        End With
    End If
    
    Set IGenericSequenceProvider_ToList = Result
    
End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As Object: Set IGenericIterator_Current = This.Iterator.Current: End Property

Private Function IGenericIterator_HasNext(Optional Out As Object) As Boolean
    
    If (This.Remaining > 0) Then
        If This.Iterator.HasNext(Out) Then
            This.Remaining = This.Remaining - 1
            IGenericIterator_HasNext = True
            Exit Function
        End If
    End If
    
    This.Remaining = 0
    Set This.Iterator = Nothing
       
End Function


