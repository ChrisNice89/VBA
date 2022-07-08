VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSequenceMapper"
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
    Mapper As GenericMapper
    Current As Object
End Type
Private This As Member

Public Property Get Source() As IGenericSequence: Set Source = This.Source: End Property
Public Property Get Mapper() As GenericMapper: Set Mapper = This.Mapper: End Property

Private Function IGenericSequence_Iterator() As IGenericIterator
   
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequence_Iterator").IsNothing("Sequence")
    
    If (This.Mapper Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequence_Iterator").IsNothing("Mapper")
    
    Set This.Iterator = This.Source.Iterator
 
    Set IGenericSequence_Iterator = Me
    
End Function

Public Function Of(ByVal Source As IGenericSequence, ByVal Mapper As GenericMapper) As IGenericSequence

    If (Source Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Of").IsNothing("Sequence")
        
    If (Mapper Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Of").IsNothing("Mapper")
    
    If Not Me Is GenericSequenceMapper Then _
        Call GenericError.Build(Me, "Public Function Of").InvalidOperation("instance call")
        
    Set This.Source = Source
    Set This.Mapper = Mapper
    
    Set Of = New GenericSequenceMapper
    This = EmptyStructure
    
End Function
'@Ignore NonReturningFunction, EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Sub Class_Initialize()

    With GenericSequenceMapper
        Set This.Source = .Source
        Set This.Mapper = .Mapper
    End With
    
End Sub

Private Function IGenericSequenceProvider_Count() As Long
    
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    
    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_Count").IsNothing("Sequence")
            
        Case TypeOf This.Source Is IGenericCollection
            Set Collection = This.Source
            IGenericSequenceProvider_Count = Collection.Count
            
        Case TypeOf This.Source Is IGenericSequenceProvider
            Set Provider = This.Source
            IGenericSequenceProvider_Count = Provider.Count
           
        Case Else
            With This.Source.Iterator
                Do While .HasNext(): IGenericSequenceProvider_Count = IGenericSequenceProvider_Count + 1: Loop
            End With
    
    End Select
    
End Function

Private Function IGenericSequenceProvider_First() As Object
    
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    '@Ignore VariableNotAssigned
    Dim Element As Object
   
    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_First").IsNothing("This.Source")
           
        Case (TypeOf This.Source Is IGenericCollection)
            Set Collection = This.Source
            If (Collection.Count > 0) Then _
                Set IGenericSequenceProvider_First = This.Mapper.Map(Collection.GetAt(0))
        
        Case TypeOf This.Source Is IGenericSequenceProvider
            Set Provider = This.Source
            Set IGenericSequenceProvider_First = This.Mapper.Map(Provider.First)
            
        Case Else
            If This.Source.Iterator.HasNext(Element) Then _
                Set IGenericSequenceProvider_First = This.Mapper.Map(Element)
            
    End Select
  
End Function

Private Function IGenericSequenceProvider_Last() As Object
    
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    '@Ignore VariableNotAssigned
    Dim Element As Object
   
    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_Last").IsNothing("This.Source")
    
        Case TypeOf This.Source Is IGenericCollection
            Set Collection = This.Source
            If Collection.Count > 0 Then _
                Set IGenericSequenceProvider_Last = This.Mapper.Map(Collection.GetAt(Collection.Count - 1))
            
        Case TypeOf This.Source Is IGenericSequenceProvider
            Set Provider = This.Source
            Set IGenericSequenceProvider_Last = This.Mapper.Map(Provider.Last)
        
        Case Else
            With This.Source.Iterator
                If .HasNext(Element) Then
                    '@Ignore EmptyDoWhileBlock
                    Do While .HasNext(Element): Loop
                    Set IGenericSequenceProvider_Last = This.Mapper.Map(Element)
                End If
            End With
            
    End Select
    
End Function

Private Function IGenericSequenceProvider_GetAt(ByVal Index As Long) As Object

    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    '@Ignore VariableNotAssigned
    Dim Element As Object
    
    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GetAt").IsNothing("This.Source")
            
        Case TypeOf This.Source Is IGenericCollection
            Set Collection = This.Source
            If Collection.Count > Index Then
                Set IGenericSequenceProvider_GetAt = This.Mapper.Map(Collection.GetAt(Index))
                Exit Function
            End If
            
        Case TypeOf This.Source Is IGenericSequenceProvider
            Set Provider = This.Source
            Set IGenericSequenceProvider_GetAt = This.Mapper.Map(Provider.GetAt(Index))
            Exit Function
           
        Case Else
            With This.Source.Iterator
                Do While .HasNext(Element)
                    If (Index = 0) Then
                        Set IGenericSequenceProvider_GetAt = This.Mapper.Map(Element)
                        Exit Function
                    End If
                    Index = Index - 1
                Loop
            End With
            
    End Select
    
    Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GetAt").IndexOutofRange
    
End Function

Private Function IGenericSequenceProvider_ToArray() As GenericArray
    
    '@Ignore VariableNotAssigned
    Dim Element As Object
    Dim Result As GenericArray, tmp As GenericArray
    
    Dim Capacity As Long, Count As Long
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_ToArray").IsNothing("This.Source")
    
    Set Result = GenericArray ' Avoid null
        
    With This.Source.Iterator
        If .HasNext(Element) Then
                    
            Capacity = IGenericSequenceProvider_GuessCount
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

                Set Result.ElementAt(Count) = This.Mapper.Map(Element)
                Count = Count + 1
                
            Loop While .HasNext(Element)

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
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_ToList").IsNothing("This.Source")
    
    Set Result = GenericOrderedList.Build
    
    With This.Source.Iterator
        If .HasNext(Element) Then
            
            Capacity = IGenericSequenceProvider_GuessCount
            If (Capacity < 0) Then _
                Capacity = CACHE_START_SIZE
                
            Call Result.EnsureCapacity(Capacity)
            
            Do: Call Result.Add(This.Mapper.Map(Element)): Loop While .HasNext(Element)
            
        End If
    End With
    
    Set IGenericSequenceProvider_ToList = Result
    
End Function

Private Function IGenericSequenceProvider_GuessCount() As Long
    
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection
    Dim guess As Long
    
    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GuessCount").IsNothing("This.Source")
            
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

'IGenericIterator
Private Property Get IGenericIterator_Current() As Object: Set IGenericIterator_Current = This.Current: End Property

Private Function IGenericIterator_HasNext(Optional Out As Object) As Boolean
 
    If This.Iterator.HasNext(Out) Then
        Set This.Current = This.Mapper.Map(Out)
        Set Out = This.Current
        IGenericIterator_HasNext = True
        Exit Function
    End If
    
    Set This.Current = Nothing
    Set This.Iterator = Nothing
    
End Function


