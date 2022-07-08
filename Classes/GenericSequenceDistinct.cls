VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericSequenceDistinct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IGenericSequenceProvider
Implements IGenericSequence
Implements IGenericIterator

Private Const CACHE_START_SIZE As Long = 2 ^ 13

Private Type Member
    Source As IGenericSequence
    Iterator As IGenericIterator
    HashProvider As IGenericHashProvider
    Cache As GenericHashSet
End Type
Private This As Member

Public Property Get Source() As IGenericSequence: Set Source = This.Source: End Property
Public Property Get HashProvider() As IGenericHashProvider: Set HashProvider = This.HashProvider: End Property

Private Function IGenericSequence_Iterator() As IGenericIterator
    
    Dim Capacity As Long
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequence_Iterator").IsNothing("This.Source")
   
    Capacity = GuessCount(This.Source)
    If (Capacity < 0) Then _
        Capacity = CACHE_START_SIZE
    
    Set This.Cache = GenericHashSet.Build(Capacity:=Capacity, HashProvider:=This.HashProvider)
    Set This.Iterator = This.Source.Iterator
    Set IGenericSequence_Iterator = Me
    
End Function

Public Function Of( _
                        ByVal Source As IGenericSequence, _
                        Optional ByVal HashProvider As IGenericHashProvider) As IGenericSequence

    If (Source Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Of").IsNothing("Sequence")
    
    If Not Me Is GenericSequenceDistinct Then _
        Call GenericError.Build(Me, "Public Function Of").InvalidOperation("instance call")
        
    Set This.Source = Source
    Set This.HashProvider = HashProvider
    
    Set Of = New GenericSequenceDistinct
    This = EmptyStructure
    
End Function
'@Ignore NonReturningFunction, EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Sub Class_Initialize()

    Set This.Source = GenericSequenceDistinct.Source
    Set This.HashProvider = GenericSequenceDistinct.HashProvider
    
End Sub

Private Function CreateCache(ByVal Sequence As IGenericSequence, ByVal HashProvider As IGenericHashProvider) As GenericHashSet

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
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_Count").IsNothing("This.Source")

    IGenericSequenceProvider_Count = CreateCache(This.Source, This.HashProvider).Elements.Count
    
End Function

Private Function IGenericSequenceProvider_GuessCount() As Long: IGenericSequenceProvider_GuessCount = GuessCount(This.Source): End Function

Private Function IGenericSequenceProvider_First() As Object
    
    Dim Provider As IGenericSequenceProvider
    Dim Collection As IGenericCollection

    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_First").IsNothing("This.Source")
           
        Case (TypeOf This.Source Is IGenericCollection)
            Set Collection = This.Source
            If (Collection.Count > 0) Then _
                Set IGenericSequenceProvider_First = Collection.GetAt(0)
        
        Case (TypeOf This.Source Is IGenericSequenceProvider)
            Set Provider = This.Source
            Set IGenericSequenceProvider_First = Provider.First
           
        Case Else
            Call This.Source.Iterator.HasNext(IGenericSequenceProvider_First)
            
    End Select
  
End Function

Private Function IGenericSequenceProvider_Last() As Object
  
    Select Case True
        Case (This.Source Is Nothing)
            Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_Last").IsNothing("This.Source")
           
        Case Else
            Set IGenericSequenceProvider_Last = CreateCache(This.Source, This.HashProvider).LastEntry 'This is an implementation Detail and could change in future -> use code below
            
    End Select
  
'    With This.Source.Iterator
'        If .HasNext(Element) Then
'
'            Capacity = IGenericSequenceProvider_GuessCount
'            If (Capacity > 0) Then
'                Set HashSet = GenericHashSet.Build(Capacity:=Capacity, Comparer:=this.HashProvider)
'            Else
'                Set HashSet = GenericHashSet.Build(Capacity:=CACHE_START_SIZE, Comparer:=this.HashProvider)
'            End If
'
'            Do
'                If HashSet.TryAdd(Element) Then _
'                    Set IGenericSequenceProvider_Last = Element
'
'            Loop While .HasNext(Element)
'
'        End If
'    End With
    
End Function

Private Function IGenericSequenceProvider_GetAt(ByVal Index As Long) As Object

    '@Ignore VariableNotAssigned
    Dim Element As Object
    Dim Cache As GenericHashSet
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GetAt").IsNothing("This.Source")

    With This.Source.Iterator
        If .HasNext(Element) Then
    
            If (Index = 0) Then
                Set IGenericSequenceProvider_GetAt = Element
                Exit Function
            End If
            
            Set Cache = GenericHashSet.Build(Capacity:=Index, HashProvider:=This.HashProvider)
            
            Do
                If Cache.TryAdd(Element) Then
                    If (Index = 0) Then
                        Set IGenericSequenceProvider_GetAt = Element
                        Exit Function
                    End If
                    Index = Index - 1
                End If
            Loop While .HasNext(Element)
        
        End If
    End With
    
    Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_GetAt").IndexOutofRange
    
End Function
'Todo Test performance
Private Function IGenericSequenceProvider_ToArray() As GenericArray
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_ToArray").IsNothing("This.Source")

    Set IGenericSequenceProvider_ToArray = CreateCache(This.Source, This.HashProvider).Elements.ToArray
    
End Function

Private Function IGenericSequenceProvider_ToList() As GenericOrderedList
    
    If (This.Source Is Nothing) Then _
        Call GenericError.Build(Me, "Private Function IGenericSequenceProvider_ToList").IsNothing("This.Source")

    Set IGenericSequenceProvider_ToList = GenericOrderedList.Build().AddAll(CreateCache(This.Source, This.HashProvider))
    
End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As Object: Set IGenericIterator_Current = This.Iterator.Current: End Property

Private Function IGenericIterator_HasNext(Optional Out As Object) As Boolean
   
    Do While This.Iterator.HasNext(Out)
        If This.Cache.TryAdd(Out) Then _
            IGenericIterator_HasNext = True: Exit Function
    Loop
    
    Set This.Cache = Nothing
    Set This.Iterator = Nothing
    
End Function


