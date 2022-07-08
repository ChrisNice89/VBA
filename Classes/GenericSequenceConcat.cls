VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSequenceConcat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IGenericSequence
Implements IGenericIterator
Implements IGenericSequenceProvider

Private Const CACHE_START_SIZE As Long = 2 ^ 12

Private Type Member
    Iterator As IGenericIterator
    Sequences() As IGenericSequence
    Index As Long
    Count As Long
End Type
Private This As Member

Public Property Get Count() As Long: Count = This.Count: End Property
Public Property Get SequenceAt(ByVal Index As Long) As IGenericSequence: Set SequenceAt = This.Sequences(Index): End Property

Private Function GuessCount(ByVal Sequence As IGenericSequence) As Long
    
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    
    If Not Sequence Is Nothing Then
        Select Case True
            Case (TypeOf Sequence Is IGenericCollection)
                Set Collection = Sequence
                GuessCount = Collection.Count
            
            Case (TypeOf Sequence Is IGenericSequenceProvider)
                Set Provider = Sequence
                GuessCount = Provider.GuessCount
               
            Case Else
                GuessCount = -1
                
        End Select
    End If
    
End Function

Private Function IGenericSequenceProvider_Count() As Long
    
    Dim i As Long
    Dim Sequence As IGenericSequence
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    
    If (This.Count = 0) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_Count").IsNothing("This.Source")

    For i = 0 To This.Count - 1
        
        Set Sequence = This.Sequences(i)
        
        If Not Sequence Is Nothing Then
            Select Case True
                Case (TypeOf Sequence Is IGenericCollection)
                    Set Collection = Sequence
                    IGenericSequenceProvider_Count = IGenericSequenceProvider_Count + Collection.Count
                
                Case (TypeOf Sequence Is IGenericSequenceProvider)
                    Set Provider = Sequence
                    IGenericSequenceProvider_Count = IGenericSequenceProvider_Count + Provider.Count
                   
                Case Else
                    With Sequence.Iterator
                        Do While .HasNext: IGenericSequenceProvider_Count = IGenericSequenceProvider_Count + 1: Loop
                    End With
            End Select
        End If
        
    Next
  
End Function

Private Function IGenericSequenceProvider_GuessCount() As Long
    
    Dim i As Long
    Dim guess As Long
    
    If (This.Count = 0) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GuessCount").IsNothing("This.Source")

    For i = 0 To This.Count - 1
        guess = GuessCount(This.Sequences(i))
        If (guess > 0) Then _
            IGenericSequenceProvider_GuessCount = IGenericSequenceProvider_GuessCount + guess
    Next

End Function

Private Function IGenericSequenceProvider_First() As Object
    
    Dim Sequence As IGenericSequence
    Dim Collection As IGenericCollection
    Dim i As Long
    
    If (This.Count = 0) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_Last").IsNothing("This.Source")
    
    For i = 0 To This.Count - 1
        
        Set Sequence = This.Sequences(i)
    
        If Not Sequence Is Nothing Then
            Select Case True
                Case (TypeOf Sequence Is IGenericCollection)
                    Set Collection = Sequence
                    If (Collection.Count > 0) Then
                        Set IGenericSequenceProvider_First = Collection.GetAt(0)
                        Exit Function
                    End If
            
                Case Else
                    If Sequence.Iterator.HasNext(IGenericSequenceProvider_First) Then _
                        Exit Function
                
            End Select
        End If
    Next
  
End Function

Private Function IGenericSequenceProvider_Last() As Object
    
    Dim i As Long, Count As Long
    
    Dim Sequence As IGenericSequence
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    
    If (This.Count = 0) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_Last").IsNothing("This.Source")
    
    For i = This.Count - 1 To 0 Step -1
    
        Set Sequence = This.Sequences(i)
        
        If Not Sequence Is Nothing Then
            Select Case True
                Case (TypeOf Sequence Is IGenericCollection)
                    Set Collection = Sequence
                    Count = Collection.Count
                    If (Count > 0) Then _
                        Set IGenericSequenceProvider_Last = Collection.GetAt(Count - 1)
                
                Case (TypeOf Sequence Is IGenericSequenceProvider)
                    Set Provider = Sequence
                    Set IGenericSequenceProvider_Last = Provider.Last
                   
                Case Else
                    With Sequence.Iterator
                        '@Ignore EmptyDoWhileBlock
                        Do While .HasNext(IGenericSequenceProvider_Last): Loop
                    End With
                
            End Select
            Exit Function
        End If
        
    Next
    
End Function

Private Function IGenericSequenceProvider_GetAt(ByVal Index As Long) As Object
    
    Dim i As Long
    '@Ignore VariableNotAssigned
    Dim Element As Object
    Dim Sequence As IGenericSequence
    Dim Collection As IGenericCollection
    
    If (This.Count = 0) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GetAt").IsNothing("This.Source")
        
    For i = 0 To This.Count - 1
        
        Set Sequence = This.Sequences(i)
        
        If Not Sequence Is Nothing Then
            If TypeOf Sequence Is IGenericCollection Then
                Set Collection = Sequence
                If (Collection.Count > Index) Then
                    Set IGenericSequenceProvider_GetAt = Collection.GetAt(Index)
                    Exit Function
                Else
                    Index = Index - Collection.Count
                End If
            Else
                With Sequence.Iterator
                    Do While .HasNext(Element)
                        If (Index = 0) Then
                            IGenericSequenceProvider_GetAt = Element
                            Exit Function
                        End If
                        Index = Index - 1
                    Loop
                End With
            End If
        End If
        
    Next
    
    Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GetAt").IndexOutofRange
    
End Function

Private Function IGenericSequenceProvider_ToArray() As GenericArray
   
    Dim i As Long
    Dim Capacity As Long, Count As Long
    Dim Sequence As IGenericSequence
    Dim Collection As IGenericCollection
    '@Ignore VariableNotAssigned
    Dim Element As Object
    
    Dim Result As GenericArray, tmp As GenericArray
    
    If (This.Count = 0) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_ToArray").IsNothing("This.Source")
        
    Set Result = GenericArray 'avoid null
 
    For i = 0 To This.Count - 1
    
        Set Sequence = This.Sequences(i)
        
        If Not Sequence Is Nothing Then
            If TypeOf Sequence Is IGenericCollection Then
                
                Set Collection = Sequence
                If (Count + Collection.Count > Capacity) Then
                
                    If (Capacity = 0) Then
                        Capacity = IGenericSequenceProvider_GuessCount

                        If (Capacity <= 0) Then _
                            Capacity = CACHE_START_SIZE
                        
                        Set Result = GenericArray.Build(Capacity)
                    Else
                        Set tmp = Result
                        Set Result = GenericArray.Build(Count + Collection.Count)
                        Call tmp.CopyTo(Target:=Result, Index:=Result.LowerBound, BeginAt:=tmp.LowerBound, Count:=Count)
                        Capacity = Result.Length
                    End If
                End If
                        
                Call Collection.CopyTo(Target:=Result, Index:=Count)
                Count = Count + Collection.Count
                
            Else
            
                With Sequence.Iterator
                    Do While .HasNext(Element)
                
                        If (Count = Capacity) Then
                        
                            If (Capacity = 0) Then
                                Capacity = IGenericSequenceProvider_GuessCount
    
                                If (Capacity <= 0) Then _
                                    Capacity = CACHE_START_SIZE
                                
                                Set Result = GenericArray.Build(Capacity)
                                
                            Else
                            
                                Set tmp = Result
                                Set Result = GenericArray.Build(Capacity * 2)
                                Call tmp.CopyTo(Target:=Result, Index:=Result.LowerBound, BeginAt:=tmp.LowerBound, Count:=Count)
                                Capacity = Result.Length
                                
                            End If
                        End If
                        
                        Set Result.ElementAt(Count) = Element
                        Count = Count + 1
                        
                    Loop
                End With
                
            End If
        End If
        
    Next
    
    If (Count < Result.Length) Then
        Set tmp = Result
        Set Result = GenericArray.Build(Count)
        Call tmp.CopyTo(Target:=Result, Index:=Result.LowerBound, BeginAt:=tmp.LowerBound, Count:=Count)
    End If
            
    Set IGenericSequenceProvider_ToArray = Result
    
End Function

Private Function IGenericSequenceProvider_ToList() As GenericOrderedList
    
    Dim i As Long
    Dim Capacity As Long
    Dim Sequence As IGenericSequence
    Dim Result As GenericOrderedList
    
    If (This.Count = 0) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_ToArray").IsNothing("This.Source")
        
    Set Result = GenericOrderedList.Build
 
    For i = 0 To This.Count - 1
    
        Set Sequence = This.Sequences(i)
        
        If Not Sequence Is Nothing Then
            
            If (Capacity = 0) Then
                Capacity = IGenericSequenceProvider_GuessCount
                
                If (Capacity < 0) Then _
                    Capacity = CACHE_START_SIZE
                    
                Call Result.EnsureCapacity(Capacity)
                
            End If
            
            Call Result.AddAll(Sequence)
            
        End If
        
    Next
            
    Set IGenericSequenceProvider_ToList = Result
   
End Function

Public Function Append(ByVal Sequence As IGenericSequence) As GenericSequenceConcat
    
    Dim i As Long
    Dim Concat As GenericSequenceConcat

    If TypeOf Sequence Is GenericSequenceConcat Then
        Set Concat = Sequence
        
        For i = 0 To Concat.Count - 1
            If UBound(This.Sequences) = This.Count Then _
                ReDim Preserve This.Sequences(This.Count + Concat.Count)
          
            Set This.Sequences(This.Count) = Concat.SequenceAt(i)
            This.Count = This.Count + 1
        Next

    Else
    
        If UBound(This.Sequences) = This.Count Then _
            ReDim Preserve This.Sequences(This.Count * 2)
       
        Set This.Sequences(This.Count) = Sequence
        This.Count = This.Count + 1
            
    End If
    
    Set Append = Me
    
End Function

Private Function IGenericSequence_Iterator() As IGenericIterator

    If (This.Count = 0) Then
        Set IGenericSequence_Iterator = IGenericSequenceProvider
        Exit Function
    End If
    
    This.Index = 0
    Set This.Iterator = This.Sequences(This.Index).Iterator
    Set IGenericSequence_Iterator = Me
    
End Function

Public Function Of(ByVal x As IGenericSequence, ByVal y As IGenericSequence) As GenericSequenceConcat
    
    
    
    If (x Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Of").IsNothing("Sequence")
        
    If (y Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Of").IsNothing("Sequence")
    
    If Not Me Is GenericSequenceConcat Then _
        Call GenericError.Build(Me, "Public Function Of").InvalidOperation("instance call")
        
    Set Of = New GenericSequenceConcat
    Call Of.Append(x).Append(y)
    
    This = EmptyStructure
    
End Function
'@Ignore NonReturningFunction, EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Sub Class_Initialize()

    If Not Me Is GenericSequenceConcat Then _
        ReDim This.Sequences(3)
    
End Sub

'IGenericIterator
Private Property Get IGenericIterator_Current() As Object: Set IGenericIterator_Current = This.Iterator.Current: End Property

Private Function IGenericIterator_HasNext(Optional Out As Object) As Boolean
    
    If (This.Iterator.HasNext(Out)) Then _
        IGenericIterator_HasNext = True: Exit Function
    
    Do While (This.Index < This.Count - 1)
        This.Index = This.Index + 1
        Set This.Iterator = This.Sequences(This.Index).Iterator
     
        If (This.Iterator.HasNext(Out)) Then _
            IGenericIterator_HasNext = True: Exit Function
        
    Loop
    
    Set This.Iterator = Nothing
    This.Index = 0
    
End Function


