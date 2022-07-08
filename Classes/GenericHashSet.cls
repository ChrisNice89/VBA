VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericHashSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "Collections"
Option Explicit

'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericHashSet
'*
'* Purpose: The GenericHashSet class is based on the model of mathematical sets and provides high-performance set operations, such as set addition (unions) and set subtraction (except).
'*          A set is a collection that contains no duplicate elements. In simple terms, the GenericHashSet class can be thought of as a map collection without values.
'*
'*          This implementation differs from a ordinary set (pseudo random order determined by hashes calculated from the elements)
'*          in that it maintains a entries list running through all of its entries.
'*          This entries list defines the iteration ordering, which is normally the order in which elements were inserted into the set (insertion-order).
'*
'*          It is dynamic in nature means the size of the set is growing according to the need.
'*          It is also sparse, with unoccupied holes in a pre-allocated array.
'*
'*          A GenericHashSet object consists of buckets that contain the index of the collection.
'*          A bucket is a virtual subgroup of elements within the set, which makes searching and retrieving easier and faster than in most collections.
'*          Each bucket is associated with a hash code, which is generated using a hash function and is based on the element.
'*          Elements must be immutable as long as they are used in the set.
'*          When an element is added to the set, the element is placed into a bucket based on the hash code of the element.
'*          Subsequent lookups of the element use the hash code to search in only one particular bucket,
'*          thus substantially reducing the number of element comparisons required to find an element.
'*
'* Notes:  The implementation uses a hash table for storage.
'*         In a set, the elements must be unique -> duplicate elements are not allowed.
'*         Elements must be immutable as long as they are used in the GenericHashSet.
'*         It provides many mathematical set operations, such as intersection, union, and (symmetric) difference (except).
'*         The capacity is the number of elements it can hold.
'*         It is a dynamic collection means the size of the set is automatically increased when the new elements are added.
'*         If a element can be null,depends on the current HashProvider
'*         The elements are not sorted.
'*         Each element in the set must provide its own hash function.
'*         A GenericHashSet does maintain the insertion order of stored entries.
'*         This map uses double hashing. There are hashsize buckets in the map, and each bucket can contain 0 or 1 element.
'*         Note that insertion order is not affected if a element is re-inserted into set. (A element e is reinserted into a set s
'*         if s.TryAdd(e) is invoked when s.contains(e) would return true immediately prior to the invocation.)
'*         TryAdd (TryRemove) method returned a boolean that indicates if it was successfully added (removed).
'*         It provides constant-time performance for the basic operations (add, contains and remove), assuming the hash function disperses elements properly among the buckets.
'*         Performance is likely to be just slightly below that of a ordinary set, due to the added expense of maintaining the entries list.
'*
'*         If NULL values are expected in the collection, a NULL comparison operator must be provided. Otherwise null pointer exceptions can occur!
'*
'*
'*======================================================================================================================

Implements IGeneric
Implements IGenericCollection
Implements IGenericSequence
Implements IGenericPredicator

#If Win64 Then
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

Private Const x0000 As Long = &H0
Private Const x7FFFFFFF As Long = &H7FFFFFFF
Private Const x80000000 As Long = &H80000000

Private Const SHRINK_TRESHOLD As Double = 0.33 '33%
Private Const NOT_SPECIFIED As Long = -1 'Dont change!
Private Const LOAD_FACTOR As Double = 0.72
Private Const MAXPRIME As Long = 2 ^ 31 - 1
Private Const HASHPRIME As Long = 101
Private Const DEFAULT_CAPACITY As Long = 8

Private Type HashItem
    Incr As Long
    Index As Long
    Code As Long
    Collision As Long
End Type
'
Private Type Entry
    HashCode As Long
    Element As Object
End Type

Private Type Member
    HashProvider As IGenericHashProvider
    Occupancy As Long
    MaxProbe As Long
    Count As Long
    Buckets() As Long
    Entries() As Entry
    Size As Long
    Capacity As Long
End Type
Private This As Member

Public Property Get Instance() As IGeneric: Set Instance = Me: End Property
Public Property Get Elements() As IGenericCollection: Set Elements = Me: End Property
Public Property Get HashProvider() As IGenericHashProvider: Set HashProvider = This.HashProvider: End Property
Public Property Get HashSetEquality() As IGenericPredicator: Set HashSetEquality = GenericHashSet: End Property
Public Property Get Capacity() As Long: Capacity = This.Capacity: End Property
Public Property Get MaxProbe() As Long: MaxProbe = This.MaxProbe: End Property

Public Property Get LoadFactor() As Double

    If (This.Capacity > x0000) Then _
        LoadFactor = VBA.Round(This.Count / This.Capacity, 2)

End Property

Public Function Build(Optional ByVal Capacity As Long, Optional ByVal HashProvider As IGenericHashProvider) As GenericHashSet
    
    Dim newSet As Member
    
    If Not (LOAD_FACTOR >= 0.1 And LOAD_FACTOR <= 1#) Then _
       Call GenericError.Build(Me, "Build").ArgOutOfRange("Not (LoadFactor >= 0.1 And LoadFactor <= 1#")
    
    If (Capacity < x0000) Then _
       Call GenericError.Build(Me, "Build").ArgOutOfRange("Capacity < 0")
    
    Set newSet.HashProvider = HashProvider
    
    If (Capacity > x0000) Then _
        Call Malloc(newSet, VBA.Int(Capacity / LOAD_FACTOR) + 1)
   
    Set Build = New GenericHashSet
    Call System.Inject(Build, VarPtr(newSet), LenB(newSet))
    
End Function
'More like an extension
Public Function BuildFrom(ByVal Sequence As IGenericSequence, Optional ByVal HashProvider As IGenericHashProvider) As GenericHashSet
    
    Dim HashSet As GenericHashSet
    Dim Collection As IGenericCollection, Capacity As Long
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function BuildFrom").IsNothing("Sequence")

    If TypeOf Sequence Is IGenericCollection Then
        Set Collection = Sequence
        
        If TypeOf Collection Is GenericHashSet Then
            If IsComparable(Me, Collection) Then
                Set BuildFrom = Collection.Copy 'Dont call trimexcess on a copy
                Exit Function
            End If
        End If
        
        Capacity = Collection.Count
                
    End If
    
    Set HashSet = GenericHashSet.Build(Capacity:=Capacity, HashProvider:=HashProvider).DoUnion(Sequence)
    
    If (HashSet.LoadFactor < SHRINK_TRESHOLD) Then _
        Call HashSet.TrimExcess
        
    Set BuildFrom = HashSet
    
End Function

Public Function Of(ByVal HashProvider As IGenericHashProvider, ParamArray IGenericValues() As Variant) As IGenericCollection

    Dim HashSet As GenericHashSet
    Dim i As Long
    
    If IsMissing(IGenericValues) Then _
        Call GenericError.Build(Me, "Public Function Of").NoSuchElement("ParamArray IGenericValues are missing")
        
    Set HashSet = GenericHashSet.Build(Capacity:=UBound(IGenericValues) + 1, HashProvider:=HashProvider)
    
    With HashSet
        For i = LBound(IGenericValues) To UBound(IGenericValues)
            
            If Not VBA.IsObject(IGenericValues(i)) Then _
                Call GenericError.Build(Me, "Public Function Of").InvalidInput("Datatype {" & TypeName$(IGenericValues(i)) & "} not allowed")
            
            If (IGenericValues(i) Is Nothing) Then _
                Call GenericError.Build(Me, "Public Function Of").InvalidInput("Element is nothing")
            
            If Not TypeOf IGenericValues(i) Is IGeneric Then _
                Call GenericError.Build(Me, "Public Function Of").InvalidInput("Not TypeOf IGenericValues(i) Is IGeneric")
                
            Call .TryAdd(Element:=IGenericValues(i))
            
        Next
    End With
    
    Set Of = HashSet
    
End Function

Public Property Get GetAt(ByVal Index As Long) As Object
    
    If (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Public Property Get GetAt").IndexOutofRange("(Index >= This.Count)")
    
    Set GetAt = This.Entries(Index).Element

End Property

Public Function Equals(ByVal Other As GenericHashSet) As Boolean: Equals = GenericHashSet.HashSetEquality(Me, Other): End Function

Public Function Union(ByVal Sequence As IGenericSequence) As GenericHashSet
    
    Dim Result As GenericHashSet
    Dim Collection As IGenericCollection
    '@Ignore VariableNotAssigned
    Dim Candidate As Object
    
    Dim guess As Long

    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Union").IsNothing("Sequence")
    
    If (This.Count = x0000) Then
        Set Union = GenericHashSet.BuildFrom(Sequence:=Sequence, HashProvider:=This.HashProvider)
        Exit Function
    End If
    
    Set Result = Me.Elements.Copy
    Set Union = Result
    
    If (Me Is Sequence) Then _
        Exit Function
    
    If TypeOf Sequence Is IGenericCollection Then
        Set Collection = Sequence
        
        guess = Collection.Count
        
        If (guess = x0000) Then _
            Exit Function
        
    End If
    
    With Sequence.Iterator
    
        If Not .HasNext(Candidate) Then _
            Exit Function

        Call Result.EnsureCapacity(guess + Result.Elements.Count)
       
        Do: Call Result.TryAdd(Candidate): Loop While .HasNext(Candidate)
        
    End With
    
    If Result.LoadFactor < SHRINK_TRESHOLD Then _
        Call Result.TrimExcess
        
    Set Union = Result

End Function

Public Function Difference(ByVal Sequence As IGenericSequence) As GenericHashSet
    
    Dim Result As GenericHashSet
    Dim Collection As IGenericCollection, Other As GenericHashSet
    '@Ignore VariableNotAssigned
    Dim Candidate As Object
    Dim i As Long
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Difference").IsNothing("Sequence")
    
    If (Me Is Sequence) Or (This.Count = x0000) Then
        Set Difference = GenericHashSet.Build(Capacity:=0, HashProvider:=This.HashProvider)
        Exit Function
    End If
    
    If TypeOf Sequence Is IGenericCollection Then
        Set Collection = Sequence
        
        If (Collection.Count = x0000) Then
             Set Difference = Me.Elements.Copy
             Exit Function
        End If
    
    End If
    
    If TypeOf Sequence Is GenericHashSet Then _
        If IsComparable(Me, Sequence) Then _
            Set Other = Sequence
    
    If (Other Is Nothing) Then
    
        Set Result = Me.Elements.Copy
        
        With Sequence.Iterator
            Do While .HasNext(Candidate): Call Result.TryRemove(Candidate): Loop
        End With
    
    Else
        
        Set Result = GenericHashSet.Build(Capacity:=This.Count, HashProvider:=This.HashProvider)
        
        With Result
            For i = 0 To This.Count - 1
                If Not (Other.Contains(This.Entries(i).Element)) Then _
                    Call .TryAdd(This.Entries(i).Element)
            Next
        End With

    End If
    
    If Result.LoadFactor < SHRINK_TRESHOLD Then _
        Call Result.TrimExcess
            
    Set Difference = Result

End Function

Public Function SymmetricDifference(ByVal Sequence As IGenericSequence) As GenericHashSet
    
    Dim Result As GenericHashSet
    Dim Collection As IGenericCollection
    Dim Candidate As Object
    
    Dim i As Long
  
    Dim guess As Long
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function SymmetricDifference").IsNothing("Sequence")
    
    If (This.Count = x0000) Then
        Set SymmetricDifference = GenericHashSet.BuildFrom(Sequence:=Sequence, HashProvider:=This.HashProvider)
        Exit Function
    End If
    
    If (Me Is Sequence) Then
        Set SymmetricDifference = GenericHashSet.Build(Capacity:=0, HashProvider:=This.HashProvider)
        Exit Function
    End If
    
    If TypeOf Sequence Is IGenericCollection Then
        Set Collection = Sequence
        
        guess = Collection.Count
        
        If (guess = x0000) Then
             Set SymmetricDifference = Me.Elements.Copy
             Exit Function
        End If
    
    End If
    
    Set Result = GenericHashSet.Build(Capacity:=guess + This.Count, HashProvider:=This.HashProvider).DoUnion(Sequence)
   
    With Result
        For i = This.Count - 1 To 0 Step -1
            Set Candidate = This.Entries(i).Element
            If Not (.TryRemove(Candidate)) Then _
                Call .TryAdd(Candidate)
        Next
    End With
  
    If Result.LoadFactor < SHRINK_TRESHOLD Then _
        Call Result.TrimExcess
            
    Set SymmetricDifference = Result

End Function

Public Function Intersection(ByVal Sequence As IGenericSequence) As GenericHashSet
    
    Dim Result As GenericHashSet
    Dim Collection As IGenericCollection, Other As GenericHashSet
    Dim Candidate As Object
    Dim i As Long
    
    Dim guess As Long
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Intersection").IsNothing("Sequence")
        
    If (This.Count = x0000) Then
        Set Intersection = GenericHashSet.Build(Capacity:=0, HashProvider:=This.HashProvider)
        Exit Function
    End If
    
    If (Me Is Sequence) Then
        Set Intersection = Me.Elements.Copy
        Exit Function
    End If
    
    Set Result = GenericHashSet.Build(Capacity:=0, HashProvider:=This.HashProvider)
    Set Intersection = Result
    
    If TypeOf Sequence Is IGenericCollection Then
        Set Collection = Sequence
        
        guess = Collection.Count
        
        If (guess = x0000) Then _
            Exit Function
        
        If (guess > This.Count) Then _
            guess = This.Count
                  
    End If
    
    If (guess = x0000) Then _
        guess = This.Count
      
    If TypeOf Sequence Is GenericHashSet Then
        If IsComparable(Me, Sequence) Then
            Set Other = Sequence
            'If other is much smaller then this, we are ready to give up the order in favor of performance
            If (Other.Elements.Count * 1.5 < This.Count) Then _
                Set Other = Nothing
        
        End If
    End If
    
    If Other Is Nothing Then 'does not preserve order!
    
        With Sequence.Iterator
        
            If Not .HasNext(Candidate) Then _
                Exit Function
          
            Call Result.EnsureCapacity(guess)
            
            Do
                If Me.Contains(Candidate) Then _
                    Call Result.TryAdd(Candidate)
            Loop While .HasNext(Candidate)
            
        End With
    
    Else
    
        Call Result.EnsureCapacity(guess)
       
        With Other
            For i = 0 To This.Count - 1
                Set Candidate = This.Entries(i).Element
                If .Contains(Candidate) Then _
                    Call Result.TryAdd(Candidate)
            Next
        End With
    
    End If
    
    If Result.LoadFactor < SHRINK_TRESHOLD Then _
        Call Result.TrimExcess
            
    Set Intersection = Result
    
End Function

Public Function DoUnion(ByVal Sequence As IGenericSequence) As GenericHashSet
    
    '@Ignore VariableNotAssigned
    Dim Candidate As Object
  
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function DoUnion").IsNothing("Sequence")
    
    Set DoUnion = Me
    
    If (Sequence Is Me) Then _
        Exit Function
    
    With Sequence.Iterator
        Do While .HasNext(Candidate): Call Me.TryAdd(Candidate): Loop
    End With
    
End Function

Public Function DoExcept(ByVal Sequence As IGenericSequence) As GenericHashSet
        
    '@Ignore VariableNotAssigned
    Dim Candidate As Object
    Dim Collection As IGenericCollection

    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function DoExecept").IsNothing("Sequence")
    
    Set DoExcept = Me
    
    If (This.Count = x0000) Then _
        Exit Function
    
    If (Sequence Is Me) Then _
        Call Me.Elements.Clear: Exit Function
    
    If TypeOf Sequence Is IGenericCollection Then
        Set Collection = Sequence
        
        If (Collection.Count = x0000) Then _
            Exit Function
         
    End If
    
    With Sequence.Iterator
        Do While .HasNext(Candidate): Call Me.TryRemove(Candidate): Loop
    End With

End Function

Public Function DoSymmetricExcept(ByVal Sequence As IGenericSequence) As GenericHashSet
        
    '@Ignore VariableNotAssigned
    Dim Candidate As Object
    Dim Collection As IGenericCollection, Other As GenericHashSet
    Dim i As Long, LastIndex As Long
    Const IS_MARKED As Byte = 1 '!
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function DoSymmetricExcept").IsNothing("Sequence")
    
    Set DoSymmetricExcept = Me
    
    If (Sequence Is Me) Then _
        Call Me.Elements.Clear: Exit Function
   
    If (This.Count = x0000) Then _
        Call Me.DoUnion(Sequence): Exit Function
   
    If TypeOf Sequence Is IGenericCollection Then
        Set Collection = Sequence
        
        If (Collection.Count = x0000) Then _
            Exit Function
    
    End If
    
    If TypeOf Sequence Is GenericHashSet Then _
        If IsComparable(Me, Sequence) Then _
            Set Other = Sequence
 
    With Sequence.Iterator
        If (Other Is Nothing) Then  'O(n+m)
            
            LastIndex = This.Count - 1
            ReDim Index(0 To LastIndex) As Byte
            
            If (This.HashProvider Is Nothing) Then
            
                Do While .HasNext(Candidate)
                    
                    i = FindIndex(Candidate)
                    If (Not i > LastIndex) Then
                        If (i < x0000) Then
                            Call Me.TryAdd(Candidate)
                        Else
                            Index(i) = IS_MARKED
                        End If
                    End If
                    
                Loop
                
            Else
            
                Do While .HasNext(Candidate)
                    
                    i = FindIndexByHashProvider(Candidate, This.HashProvider)
                    If (Not i > LastIndex) Then
                        If (i < x0000) Then
                            Call Me.TryAdd(Candidate)
                        Else
                            Index(i) = IS_MARKED
                        End If
                    End If
                    
                Loop
                
            End If
        
            For i = LastIndex To LBound(Index) Step -1
                If (Index(i) = IS_MARKED) Then _
                    Call RemoveEntry(i)
            Next
        
        Else
    
            Do While .HasNext(Candidate)
                If Not Me.TryRemove(Candidate) Then _
                    Call Me.TryAdd(Candidate)
            Loop
        
        End If
    End With
    
End Function

Public Function DoIntersect(ByVal Sequence As IGenericSequence) As GenericHashSet

    '@Ignore VariableNotAssigned
    Dim Candidate As Object
    Dim Collection As IGenericCollection, Other As GenericHashSet
    Dim i As Long
    Const IS_MARKED As Byte = 1 '!
  
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function DoIntersect").IsNothing("Sequence")
    
    Set DoIntersect = Me
    
    If (This.Count = x0000) Then _
        Exit Function
        
    If (Sequence Is Me) Then _
        Exit Function
  
    If TypeOf Sequence Is IGenericCollection Then
        Set Collection = Sequence
        
        If (Collection.Count = x0000) Then _
             Call Me.Elements.Clear: Exit Function
        
    End If
    
    If TypeOf Sequence Is GenericHashSet Then _
        If IsComparable(Me, Sequence) Then _
            Set Other = Sequence
                
    'O(n+m)
    If (Other Is Nothing) Then
        
        With Sequence.Iterator
            
            If Not .HasNext(Candidate) Then _
                Exit Function
            
            ReDim Index(0 To This.Count - 1) As Byte
            
            If (This.HashProvider Is Nothing) Then
            
                Do
                    i = FindIndex(Candidate)
                    If (i >= x0000) Then _
                        Index(i) = IS_MARKED
                Loop While .HasNext(Candidate)
            
            Else
            
                Do
                    i = FindIndexByHashProvider(Candidate, This.HashProvider)
                    If (i >= x0000) Then _
                        Index(i) = IS_MARKED
                Loop While .HasNext(Candidate)
                
            End If
            
            For i = UBound(Index) To LBound(Index) Step -1
                If Not (Index(i) = IS_MARKED) Then _
                    Call RemoveEntry(i)
            Next
            
        End With
        
    Else
          
        With Other
            For i = This.Count - 1 To 0 Step -1
                If Not .Contains(This.Entries(i).Element) Then _
                    Call RemoveEntry(i)
            Next
        End With
        
    End If
    
End Function

Public Function Overlaps(ByVal Sequence As IGenericSequence) As Boolean
    
    '@Ignore VariableNotAssigned
    Dim Candidate As Object
    Dim Collection As IGenericCollection, Other As GenericHashSet
    Dim i As Long
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Overlaps").IsNothing("Sequence")
    
    If (This.Count = x0000) Then _
        Exit Function
    
    If (Sequence Is Me) Then _
        Overlaps = True: Exit Function
        
    If TypeOf Sequence Is IGenericCollection Then
        Set Collection = Sequence
        
        If (Collection.Count = x0000) Then _
            Exit Function
       
    End If
    
    If TypeOf Sequence Is GenericHashSet Then _
        If IsComparable(Me, Sequence) Then _
            Set Other = Sequence
            
    If (Other Is Nothing) Then
    
        With Sequence.Iterator
        
            If (This.HashProvider Is Nothing) Then
            
                Do While .HasNext(Candidate)
                    If (FindIndex(Candidate) >= 0) Then _
                        Overlaps = True: Exit Function
                Loop
                
            Else
                
                Do While .HasNext(Candidate)
                    If (FindIndexByHashProvider(Candidate, This.HashProvider) >= 0) Then _
                        Overlaps = True: Exit Function
                Loop
            
            End If
            
        End With
        
    Else
    
        With Other
            For i = 0 To This.Count - 1
                If .Contains(This.Entries(i).Element) Then _
                    Overlaps = True: Exit Function
            Next
        End With
        
    End If
    
End Function

Public Function TryAdd(ByVal Element As Object) As Boolean

    Dim Hash As HashItem
    Dim i As Long, Slot As Long
    Dim ReuseSlot As Boolean, UnusedSlot As Long
    
    Dim Casted As IGenericValue
    
    If (This.Count = This.Capacity) Then _
        Call Expand
    
    If (This.HashProvider Is Nothing) Then
    
        Set Casted = Element 'Try cast
        
        Hash.Code = (Casted.HashCode And x7FFFFFFF)
        Hash.Index = (Hash.Code Mod This.Size)
    
        i = This.Buckets(Hash.Index)
        If Not (i = x0000) Then
            
            If Not (i = x80000000) Then
                Slot = (i - 1 And x7FFFFFFF)
                If (This.Entries(Slot).HashCode = Hash.Code) Then _
                    If Casted.Equals(This.Entries(Slot).Element) Then _
                        Exit Function
                This.Buckets(Hash.Index) = (i Or x80000000)
            Else
                UnusedSlot = Hash.Index
                ReuseSlot = True
            End If
            
            Hash.Collision = 1
            Hash.Incr = (1 + (Hash.Code Mod (This.Size - 1))) 'Hash.Incr = 1 + Modulo((CDbl(.Code) * HASHPRIME), (.Size - 1))
            
            Do
        
                Hash.Index = ((Hash.Index + Hash.Incr) Mod This.Size)
                
                i = This.Buckets(Hash.Index)
                If (i = x0000) Then _
                    Exit Do
                    
                If Not (i = x80000000) Then
                    Slot = (i - 1 And x7FFFFFFF)
                    If (This.Entries(Slot).HashCode = Hash.Code) Then _
                        If Casted.Equals(This.Entries(Slot).Element) Then _
                            Exit Function
                    If Not ReuseSlot Then _
                        This.Buckets(Hash.Index) = (i Or x80000000)
                    
                ElseIf Not ReuseSlot Then
                    UnusedSlot = Hash.Index
                    ReuseSlot = True
                End If
                
            Loop While (InterlockedIncrement(Hash.Collision) < This.Size)
        
        End If
        
    Else
    
        Hash.Code = (This.HashProvider.HashOf(Element) And x7FFFFFFF)
        Hash.Index = (Hash.Code Mod This.Size)
    
        i = This.Buckets(Hash.Index)
        If Not (i = x0000) Then
            
            If Not (i = x80000000) Then
                Slot = (i - 1 And x7FFFFFFF)
                If (This.Entries(Slot).HashCode = Hash.Code) Then _
                    If This.HashProvider.Equality(This.Entries(Slot).Element, Element) Then _
                        Exit Function
                This.Buckets(Hash.Index) = (i Or x80000000)
            Else
                ReuseSlot = Hash.Index
                ReuseSlot = True
            End If
            
            Hash.Collision = 1
            Hash.Incr = (1 + (Hash.Code Mod (This.Size - 1))) 'Hash.Incr = 1 + Modulo((CDbl(.Code) * HASHPRIME), (.Size - 1))
            
            Do
            
                Hash.Index = ((Hash.Index + Hash.Incr) Mod This.Size)
                
                i = This.Buckets(Hash.Index)
                If (i = x0000) Then _
                    Exit Do
                    
                If Not (i = x80000000) Then
                    Slot = (i - 1 And x7FFFFFFF)
                    If (This.Entries(Slot).HashCode = Hash.Code) Then _
                        If This.HashProvider.Equality(This.Entries(Slot).Element, Element) Then _
                            Exit Function
                    If Not ReuseSlot Then _
                        This.Buckets(Hash.Index) = (i Or x80000000)
                    
                ElseIf Not ReuseSlot Then
                    UnusedSlot = Hash.Index
                    ReuseSlot = True
                End If
                
            Loop While (InterlockedIncrement(Hash.Collision) < This.Size)
        
        End If
        
    End If
    
    This.Entries(This.Count).HashCode = Hash.Code
    Set This.Entries(This.Count).Element = Element
    This.Count = This.Count + 1
    
    If Not (ReuseSlot) Then
        If (Hash.Collision > This.MaxProbe) Then _
            This.MaxProbe = Hash.Collision
        This.Buckets(Hash.Index) = This.Count
    Else
        This.Buckets(UnusedSlot) = (This.Count Or x80000000)
    End If
    
    TryAdd = True
    
End Function

Public Function TryRemove(ByVal Element As Object) As Boolean
    
    Dim Hash As HashItem
    Dim i As Long, Slot As Long
    Dim Entry As Entry
    
    Dim Casted As IGenericValue
    
    If (This.Count = x0000) Then _
        Exit Function
        
    If (This.HashProvider Is Nothing) Then
        
        Set Casted = Element 'Try cast
        Hash.Code = (Casted.HashCode And x7FFFFFFF)
        Hash.Index = (Hash.Code Mod This.Size)
        
        i = This.Buckets(Hash.Index)
        
        If (i = x0000) Then _
            Exit Function
            
        If Not (i = x80000000) Then
            Slot = (i - 1 And x7FFFFFFF)
            If (This.Entries(Slot).HashCode = Hash.Code) Then _
                If Casted.Equals(This.Entries(Slot).Element) Then _
                    GoTo Remove
                    
            If ((i And x80000000) = x0000) Then _
                Exit Function
        
        End If
        
        Hash.Incr = (1 + (Hash.Code Mod (This.Size - 1))) 'Hash.Incr = 1 + Modulo((CDbl(.Code) * HASHPRIME), (.Size - 1))
        
        Do
            
            Hash.Index = ((Hash.Index + Hash.Incr) Mod This.Size)
            
            i = This.Buckets(Hash.Index)
            
            If (i = x0000) Then _
                Exit Function
                
            If Not (i = x80000000) Then
                Slot = (i - 1 And x7FFFFFFF)
                If (This.Entries(Slot).HashCode = Hash.Code) Then _
                    If Casted.Equals(This.Entries(Slot).Element) Then _
                        Exit Do
            End If
            
            If ((i And x80000000) = x0000) Then _
                Exit Function
           
        Loop While (InterlockedIncrement(Hash.Collision) < This.MaxProbe)
        
    Else
    
        Hash.Code = (This.HashProvider.HashOf(Element) And x7FFFFFFF)
        Hash.Index = (Hash.Code Mod This.Size)
        
        i = This.Buckets(Hash.Index)
        
        If (i = x0000) Then _
            Exit Function
            
        If Not (i = x80000000) Then
            Slot = (i - 1 And x7FFFFFFF)
            If (This.Entries(Slot).HashCode = Hash.Code) Then _
                If This.HashProvider.Equality(This.Entries(Slot).Element, Element) Then _
                    GoTo Remove
                    
            If ((i And x80000000) = x0000) Then _
                Exit Function
        
        End If
        
        Hash.Incr = (1 + (Hash.Code Mod (This.Size - 1))) 'Hash.Incr = 1 + Modulo((CDbl(.Code) * HASHPRIME), (.Size - 1))
        
        Do
        
            Hash.Index = ((Hash.Index + Hash.Incr) Mod This.Size)
            
            i = This.Buckets(Hash.Index)
            
            If (i = x0000) Then _
                Exit Function
                
            If Not (i = x80000000) Then
                Slot = (i - 1 And x7FFFFFFF)
                If (This.Entries(Slot).HashCode = Hash.Code) Then _
                    If This.HashProvider.Equality(This.Entries(Slot).Element, Element) Then _
                        Exit Do
            End If
            
            If ((i And x80000000) = x0000) Then _
                Exit Function
            
        Loop While (InterlockedIncrement(Hash.Collision) < This.MaxProbe)
    
    End If

Remove:
    
    If (i > x0000) Then
        This.Buckets(Hash.Index) = x0000
    Else
        This.Buckets(Hash.Index) = x80000000
    End If
    
    If Not ((i And x7FFFFFFF) = This.Count) Then
    
        Entry = This.Entries(This.Count - 1)
        
        Hash.Code = Entry.HashCode
        Hash.Index = (Hash.Code Mod This.Size)
        i = This.Buckets(Hash.Index)
        
        If Not ((i And x7FFFFFFF) = This.Count) Then
            
            Hash.Incr = 1 + (Hash.Code Mod (This.Size - 1))
                
            Do
                Hash.Index = (Hash.Index + Hash.Incr) Mod This.Size
                i = This.Buckets(Hash.Index)
            Loop While Not ((i And x7FFFFFFF) = This.Count)
    
        End If
        
        This.Entries(Slot) = Entry
        
        Slot = Slot + 1 '!
        
        If (i < x0000) Then _
            Slot = (Slot Or x80000000)
        
        This.Buckets(Hash.Index) = Slot
    
    End If
    
    This.Count = This.Count - 1
    This.Entries(This.Count) = EmptyEntry
    
    TryRemove = True
    
End Function

Public Sub RemoveAt(ByVal Index As Long)

    If (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Public Sub RemoveAt").IndexOutofRange("(Index >= This.Count)")
    
    Call RemoveEntry(Index)
    
End Sub

Public Function Pop() As Object
 
    If (This.Count = x0000) Then _
        Exit Function

    Set Pop = Me.TakeAt(This.Count - 1)
     
End Function

Public Function TakeAt(ByVal Index As Long) As GenericPair

    If (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Public Sub TakeAt").IndexOutofRange("(Index >= This.Count)")
    
    Set TakeAt = This.Entries(Index).Element
    Call RemoveEntry(Index)
    
End Function

Public Function Contains(ByVal Element As Object) As Boolean: Contains = Me.IndexOf(Element) >= x0000: End Function

Public Function ContainsAll(ByVal Sequence As IGenericSequence) As Boolean
    
    '@Ignore VariableNotAssigned
    Dim Candidate As Object
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function ContainsAll").IsNothing("Sequence")
    
    If (This.Count = x0000) Then _
        Exit Function
        
    If Not (Sequence Is Me) Then
        With Sequence.Iterator
        
            If (This.HashProvider Is Nothing) Then
            
                Do While .HasNext(Candidate)
                    If (FindIndex(Candidate) < 0) Then _
                        Exit Function
                Loop
                
            Else
            
                Do While .HasNext(Candidate)
                    If (FindIndexByHashProvider(Candidate, This.HashProvider) < 0) Then _
                        Exit Function
                Loop
                
            End If
            
        End With
    End If
    
    ContainsAll = True

End Function

Public Function IndexOf(ByVal Element As Object) As Long
   
    If (This.Count > x0000) Then
        If (This.HashProvider Is Nothing) Then
            IndexOf = FindIndex(Element)
        Else
            IndexOf = FindIndexByHashProvider(Element, This.HashProvider)
        End If
        Exit Function
    End If
    
    IndexOf = -1
    
End Function

Private Sub RemoveEntry(ByVal Slot As Long)
    
    Dim Entry As Entry
    Dim Hash As HashItem
    Dim i As Long, Bucket As Long
    
    Hash.Code = This.Entries(Slot).HashCode
    Hash.Index = (Hash.Code Mod This.Size)
    
    Bucket = Slot + 1
    
    i = This.Buckets(Hash.Index)
    If Not ((i And x7FFFFFFF) = Bucket) Then
        
        Hash.Incr = 1 + (Hash.Code Mod (This.Size - 1))
            
        Do
            Hash.Index = (Hash.Index + Hash.Incr) Mod This.Size
            i = This.Buckets(Hash.Index)
        Loop While Not ((i And x7FFFFFFF) = Bucket)

    End If
    
    If (i > x0000) Then
        This.Buckets(Hash.Index) = x0000
    Else
        This.Buckets(Hash.Index) = x80000000
    End If
    
    If Not (Bucket = This.Count) Then
        
        Entry = This.Entries(This.Count - 1)
        
        Hash.Code = Entry.HashCode
        Hash.Index = (Hash.Code Mod This.Size)
        
        i = This.Buckets(Hash.Index)
        If Not ((i And x7FFFFFFF) = This.Count) Then
            
            Hash.Incr = 1 + (Hash.Code Mod (This.Size - 1))
                
            Do
                Hash.Index = (Hash.Index + Hash.Incr) Mod This.Size
                i = This.Buckets(Hash.Index)
            Loop While Not ((i And x7FFFFFFF) = This.Count)
    
        End If
        
        If (i < x0000) Then _
            Bucket = (Bucket Or x80000000)
        
        This.Buckets(Hash.Index) = Bucket
        This.Entries(Slot) = Entry
    
    End If

    This.Count = This.Count - 1
    This.Entries(This.Count) = EmptyEntry
    
End Sub

Public Function OccupiedSlots() As Long

    Dim i As Long
    For i = LBound(This.Buckets) To UBound(This.Buckets)
        If This.Buckets(i) < x0000 Then
            OccupiedSlots = OccupiedSlots + 1
        End If
    Next

End Function

Public Function Slice(ByVal BeginAt As Long, Optional ByVal Count As Long = NOT_SPECIFIED) As GenericHashSet
    
    Dim SubSet As GenericHashSet
    Dim i As Long
    
    If (BeginAt < x0000) Then _
        Call GenericError.Build(Me, "Public Function Slice").ArgOutOfRange("(BeginAt < 0)")
        
    If (Count = NOT_SPECIFIED) Then _
        Count = This.Count - BeginAt
   
    If (Count <= x0000) Then _
        Call GenericError.Build(Me, "Public Function Slice").ArgOutOfRange("(Count <= 0)")
        
    If (This.Count - BeginAt < Count) Then _
        Call GenericError.Build(Me, "Public Function Slice").InvalidInput("(This.Count - BeginAt < Count)")
   
    Set SubSet = GenericHashSet.Build(Capacity:=Count, HashProvider:=This.HashProvider)
    
    With SubSet
        For i = BeginAt To BeginAt + Count - 1: Call .TryAdd(This.Entries(i).Element): Next
    End With
    
    Set Slice = SubSet
    
End Function

Public Function LastEntry() As Object
    
    If (This.Count = x0000) Then _
        Exit Function
    
    Set LastEntry = This.Entries(This.Count - 1).Element
    
End Function

Public Function Reverse() As IGenericIterator: Set Reverse = GenericIndexIterator.Build(Me, 0, This.Count, True): End Function

Public Function Range(Optional ByVal BeginAt As Long = NOT_SPECIFIED, Optional ByVal Count As Long = NOT_SPECIFIED) As IGenericIterator
    
    If (BeginAt = NOT_SPECIFIED) Then
        BeginAt = 0
    ElseIf (BeginAt < 0) Then
        Call GenericError.Build(Me, "Public Function Range").ArgOutOfRange("BeginAt < 0")
    End If
    
    If (Count = NOT_SPECIFIED) Then
        Count = This.Count - BeginAt
    ElseIf (Count < 0) Then
        Call GenericError.Build(Me, "Public Function Range").ArgOutOfRange("Count < 0")
    End If
    
    If (This.Count - BeginAt < Count) Then _
        Call GenericError.Build(Me, "Public Function Range").InvalidInput("(This.Count - BeginAt < Count)")
   
    Set Range = GenericIndexIterator.Build(Me, BeginAt, Count)
    
End Function

'Allocates new bucket array with prime length
Private Function Malloc(ByRef Struct As Member, ByVal Size As Long) As Long
    
    If (Size <= x0000) Then _
        Call GenericError.Build(Me, "Private Function Malloc").InvalidInput("(Size < x0000)")
        
    If (Size < This.Count) Then _
        Call GenericError.Build(Me, "Private Function Malloc").InvalidInput("(Size < This.Count)")
        
    If (Size > GenericArray.MaxLength) Then _
        Call GenericError.Build(Me, "Private Function Malloc").Overflow("Size > GenericArray.MaxLength")
    
    Size = GetPrime(Size)

    Struct.Capacity = VBA.Int(Size * LOAD_FACTOR)
    
    If Not (Struct.Capacity < Size) Then _
        Call GenericError.Build(Me, "Private Function Malloc").InvalidInput("Based on the current algorithm, Capacity must be less than Size.")
        
    Struct.Size = Size
    Struct.Occupancy = x0000
    
    If (Struct.Count > x0000) Then
        ReDim Preserve Struct.Entries(Struct.Capacity - 1)
    Else
        ReDim Struct.Entries(Struct.Capacity - 1)
    End If
    
    ReDim Struct.Buckets(Struct.Size - 1)
    
    Malloc = Size
    
End Function

'@Ignore EmptyMethod
Private Function EmptyEntry() As Entry: End Function

Public Function TrimExcess() As GenericHashSet
    
    Dim Size As Long
    
    If (Me Is GenericHashSet) Then _
        Call GenericError.Build(Me, "Public Function TrimExcess").InvalidOperation("(Me Is GenericHashSet)")
        
    If (This.Capacity > x0000) Then
            
        Size = VBA.Int(This.Count / LOAD_FACTOR) + 1
        
        If (GetPrime(Size) < This.Size) Then _
            Call Resize(Size)
  
    End If
    
    Set TrimExcess = Me
    
End Function

Public Function EnsureCapacity(ByVal Min As Long) As GenericHashSet
    
    Dim Size As Long
    
    If (Me Is GenericHashSet) Then _
        Call GenericError.Build(Me, "Public Function EnsureCapacity").InvalidOperation("(Me Is GenericHashSet)")
        
    Size = VBA.Int(Min / LOAD_FACTOR) + 1
    
    If (Size > This.Capacity) Then _
        Call Resize(Size)
    
    Set EnsureCapacity = Me
    
End Function

Private Sub Expand()
    
    Dim Size As Long
    
    If (Me Is GenericHashSet) Then _
        Call GenericError.Build(Me, "Private Sub Expand").InvalidOperation("(Me Is GenericHashSet)")
        
    If (This.Size > x0000) Then
        Size = This.Size * 2
    Else
        Size = VBA.Int(DEFAULT_CAPACITY / LOAD_FACTOR) + 1
    End If
    
    Call Resize(Size)
    
End Sub

Public Function Rehash(Optional ByVal ForceRehashing As Boolean = False) As GenericHashSet
    
    Dim i As Long
    
    Set Rehash = Me

    If (This.Capacity = x0000) Then _
        Exit Function

    If ForceRehashing Then
        If Not (This.HashProvider Is Nothing) Then
            With This.HashProvider
                For i = 0 To This.Count - 1: This.Entries(i).HashCode = (.HashOf(This.Entries(i).Element) And x7FFFFFFF): Next
            End With
        End If
    End If

    ReDim This.Buckets(This.Size - 1)
    Call LoadBuckets(This.Buckets, This.Entries)
    
End Function

Private Sub Resize(ByVal Size As Long)
   
    'Clear cache and allocate new
    Call Malloc(This, Size)
    
    If (This.Count = x0000) Then _
        Exit Sub
    
    Call LoadBuckets(This.Buckets, This.Entries)
                                                                                            
End Sub

Private Sub LoadBuckets(ByRef Buckets() As Long, ByRef Entries() As Entry)

    Dim Hash As HashItem
    Dim i As Long, Slot As Long

    This.MaxProbe = x0000
    This.Occupancy = x0000

    For i = 1 To This.Count

        Hash.Code = This.Entries(i - 1).HashCode
        Hash.Index = (Hash.Code Mod This.Size)
        Slot = This.Buckets(Hash.Index)

        If Not (Slot = x0000) Then

            Hash.Incr = 1 + (Hash.Code Mod (This.Size - 1))

            Do
                Hash.Collision = Hash.Collision + 1
                This.Buckets(Hash.Index) = (Slot Or x80000000)
                Hash.Index = (Hash.Index + Hash.Incr) Mod This.Size
                Slot = This.Buckets(Hash.Index)
            Loop While Not (Slot = x0000)

            If (Hash.Collision > This.MaxProbe) Then _
                This.MaxProbe = Hash.Collision
            Hash.Collision = x0000

        End If

        This.Buckets(Hash.Index) = i

    Next

End Sub

'TODO Check HashProvider Equality
Private Function IsComparable(ByVal A As GenericHashSet, ByVal B As GenericHashSet) As Boolean
    IsComparable = (A.HashProvider Is B.HashProvider)
End Function

Private Function GetPrime(Optional ByVal Min As Long = 0) As Long
    
    Const LAST_INDEX As Long = 71
    Static Primes() As Long
    Dim i As Long
    Dim Prime As Variant
    
    If (Min < x0000) Then _
        Call GenericError.Build(Me, "getPrime").InvalidInput("Min < 0")
    
    If ((Not Primes) = -1) Then
        ReDim Primes(LAST_INDEX)
        For Each Prime In Array(3, 7, 11, 17, 23, 29, 37, 47, 59, 71, 89, 107, 131, 163, 197, 239, 293, 353, 431, 521, 631, 761, 919, _
                            1103, 1327, 1597, 1931, 2333, 2801, 3371, 4049, 4861, 5839, 7013, 8419, 10103, 12143, 14591, _
                            17519, 21023, 25229, 30293, 36353, 43627, 52361, 62851, 75431, 90523, 108631, 130363, 156437, _
                            187751, 225307, 270371, 324449, 389357, 467237, 560689, 672827, 807403, 968897, 1162687, 1395263, _
                            1674319, 2009191, 2411033, 2893249, 3471899, 4166287, 4999559, 5999471, 7199369)

            Primes(i) = Prime
            i = i + 1
        Next
    End If

    For i = 3 To UBound(Primes)
        GetPrime = Primes(i)
        If (GetPrime >= Min) Then _
            Exit Function
    
    Next
   
    Do While Min <= MAXPRIME
        If IsPrime(Min) Then _
            If (((Min - 1) Mod HASHPRIME) = x0000 = False) Then _
                Exit Do
        Min = Min + 1
    Loop
   
    GetPrime = Min
    
End Function

Private Function IsPrime(ByVal Candidate As Long) As Boolean
    
    Dim Divisor As Long
 
    If Not ((Candidate And 1) = x0000) Then
        
        For Divisor = 3 To CLng(VBA.Sqr(Candidate)) Step 2
            If ((Candidate Mod Divisor) = x0000) Then _
                Exit Function
        Next
        
        IsPrime = True
        Exit Function
        
    End If
    
    IsPrime = (Candidate = 2)
    
End Function
'Try cast
Private Function FindIndex(ByVal Element As IGenericValue) As Long

    Dim Hash As HashItem
    Dim i As Long
    
    Hash.Code = (Element.HashCode And x7FFFFFFF)
    Hash.Index = (Hash.Code Mod This.Size)
    
    i = This.Buckets(Hash.Index)
    
    If (i = x0000) Then _
        Exit Function
        
    If Not (i = x80000000) Then
        FindIndex = (i - 1 And x7FFFFFFF)
        
        If (This.Entries(FindIndex).HashCode = Hash.Code) Then _
            If (Element.Equals(This.Entries(FindIndex).Element)) Then _
                Exit Function
        
        If ((i And x80000000) = x0000) Then _
            Exit Function
                    
    End If
        
    Hash.Incr = (1 + (Hash.Code Mod (This.Size - 1))) 'Hash.Incr = 1 + Modulo((CDbl(.Code) * HASHPRIME), (.Size - 1))
    
    Do
        
        Hash.Index = (Hash.Index + Hash.Incr) Mod This.Size
        i = This.Buckets(Hash.Index)
        
        If (i = x0000) Then _
            Exit Do
            
        If Not (i = x80000000) Then
            FindIndex = (i - 1 And x7FFFFFFF)
            
            If (This.Entries(FindIndex).HashCode = Hash.Code) Then _
                If (Element.Equals(This.Entries(FindIndex).Element)) Then _
                    Exit Function
            
            '(Bucket >= 0)
            If ((i And x80000000) = x0000) Then _
                Exit Do
                
        End If

    Loop While (InterlockedIncrement(Hash.Collision) < This.MaxProbe)

    FindIndex = -1
    
End Function

Private Function FindIndexByHashProvider(ByVal Element As Object, ByVal HashProvider As IGenericHashProvider) As Long

    Dim Hash As HashItem
    Dim i As Long
    
    Hash.Code = (HashProvider.HashOf(Element) And x7FFFFFFF)
    Hash.Index = (Hash.Code Mod This.Size)
    
    i = This.Buckets(Hash.Index)
    
    If (i = x0000) Then _
        Exit Function
        
    If Not (i = x80000000) Then
        FindIndexByHashProvider = (i - 1 And x7FFFFFFF)
        
        If (This.Entries(FindIndexByHashProvider).HashCode = Hash.Code) Then _
            If (HashProvider.Equality(This.Entries(FindIndexByHashProvider).Element, Element)) Then _
                Exit Function
                
        If ((i And x80000000) = x0000) Then _
            Exit Function
            
    End If

    Hash.Incr = (1 + (Hash.Code Mod (This.Size - 1))) 'Hash.Incr = 1 + Modulo((CDbl(.Code) * HASHPRIME), (.Size - 1))
    
    Do
        
        Hash.Index = (Hash.Index + Hash.Incr) Mod This.Size
        i = This.Buckets(Hash.Index)
        
        If (i = 0) Then _
            Exit Do
            
        If Not (i = x80000000) Then
            FindIndexByHashProvider = (i - 1 And x7FFFFFFF)
            
            If (This.Entries(FindIndexByHashProvider).HashCode = Hash.Code) Then _
                If (HashProvider.Equality(This.Entries(FindIndexByHashProvider).Element, Element)) Then _
                    Exit Function
            
            '(Bucket >= 0)
            If ((i And x80000000) = x0000) Then _
                Exit Do
                
        End If

    Loop While (InterlockedIncrement(Hash.Collision) < This.MaxProbe)
    
    FindIndexByHashProvider = -1
    
End Function

Private Function Modulo(ByVal A As Double, ByVal m As Long) As Long: Modulo = (A - (VBA.Int(A / m) * m)): End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericHashSet: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = GenericHashSet: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(This): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(This): End Property
Private Property Get IGeneric_HashCode() As Long:  IGeneric_HashCode = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericHashSet):  End Property

'@Ignore EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As Object) As Boolean

    If (Other Is Nothing) Then _
        Exit Function
    
    If TypeOf Other Is GenericHashSet Then _
        IGeneric_Equals = Me.Equals(Other)
    
End Function

Private Function IGeneric_ToString() As String
    With This
        IGeneric_ToString = _
                            "Class: " & VBA.TypeName$(Me) & vbNewLine & _
                            "Size: " & .Size & vbNewLine & _
                            "Capacity: " & .Capacity & vbNewLine & _
                            "LoadFactor: " & Me.LoadFactor & vbNewLine & _
                            "Elements: " & .Count & vbNewLine & _
                            "Occupancy: " & .Occupancy
    End With
End Function

'IGenericCollection
Private Property Get IGenericCollection_HasRandomAccess() As Boolean: IGenericCollection_HasRandomAccess = True: End Property
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = This.Count: End Property

Private Function IGenericCollection_GetAt(ByVal Index As Long) As Object

    If (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Private Function IGenericCollection_GetAt").IndexOutofRange("(Index >= This.Count)")
    
    Set IGenericCollection_GetAt = This.Entries(Index).Element

End Function

Private Function IGenericCollection_Contains(ByVal Element As Object) As Boolean

    If (This.Count = x0000) Then _
        Exit Function
    
    If (This.HashProvider Is Nothing) Then
        IGenericCollection_Contains = FindIndex(Element) >= 0
    Else
        IGenericCollection_Contains = FindIndexByHashProvider(Element, This.HashProvider) >= 0
    End If

End Function

Private Sub IGenericCollection_CopyTo(ByVal Target As GenericArray, ByVal Index As Long)
    
    Dim i As Long
    
    If (Target Is Nothing) Then _
        Call GenericError.Build(Me, "Private Sub IGenericCollection_CopyTo").IsNothing("Target")
    
    If (Index < Target.LowerBound) Or (Index >= Target.Length) Then _
        Call GenericError.Build(Me, "Private Sub IGenericCollection_CopyTo").IndexOutofRange("(Index < Target.LowerBound) Or (Index >= Target.Length)")
    
    If (This.Count > Target.Length - Index) Then _
        Call GenericError.Build(Me, "Private Sub IGenericCollection_CopyTo").InvalidInput("(This.Count > Target.length - Index)")
        
    If (This.Count = x0000) Then _
        Exit Sub
        
    With Target
        For i = 0 To This.Count - 1: Set .ElementAt(Index + i) = This.Entries(i).Element: Next
    End With
    
End Sub

Private Sub IGenericCollection_Clear()

    If (This.Capacity = x0000) Then _
        Exit Sub
        
    ReDim This.Buckets(This.Size - 1)
    ReDim This.Entries(This.Capacity - 1)
    
    This.Count = x0000
    This.Occupancy = x0000
    
End Sub

Private Function IGenericCollection_ToArray() As GenericArray
    
    Set IGenericCollection_ToArray = GenericArray 'avoid null
    
    If (This.Count = x0000) Then _
        Exit Function
        
    Set IGenericCollection_ToArray = GenericArray.Build(This.Count)
    Call Me.Elements.CopyTo(IGenericCollection_ToArray, IGenericCollection_ToArray.LowerBound)
    
End Function

Private Function IGenericCollection_Copy() As IGenericCollection
    
    Dim Clone As Member
    With Clone
        Set .HashProvider = This.HashProvider
        .Capacity = This.Capacity
        .Count = This.Count
        .Size = This.Size
        .Occupancy = This.Occupancy
        .MaxProbe = This.MaxProbe
        .Buckets = This.Buckets
        .Entries = This.Entries
    End With
    
    Set IGenericCollection_Copy = New GenericHashSet
    Call System.Inject(IGenericCollection_Copy, VarPtr(Clone), LenB(Clone))
    
End Function

Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Range: End Function

'IGenericSequence_Iterator
Private Function IGenericSequence_Iterator() As IGenericIterator: Set IGenericSequence_Iterator = Me.Range: End Function

'IGenericPredicator
Private Function IGenericPredicator_IsTrue(ByVal A As Object, ByVal B As Object) As Boolean
    
    Dim First As GenericHashSet, Second As GenericHashSet
    
    If A Is B Then _
        IGenericPredicator_IsTrue = True: Exit Function
        
    If A Is Nothing Then _
        Exit Function
    
    If B Is Nothing Then _
        Exit Function
    
    Set First = A: Set Second = B 'Try cast
    
    If Not IsComparable(First, Second) Then _
        Exit Function
        
    If Not (First.Elements.Count = Second.Elements.Count) Then _
        Exit Function
    
    IGenericPredicator_IsTrue = First.ContainsAll(Second)
    
End Function


