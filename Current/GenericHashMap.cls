VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericHashMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@PredeclaredId
'@IgnoreModule FunctionReturnValueAlwaysDiscarded, VariableNotAssigned, NonReturningFunction, UnassignedVariableUsage, ProcedureNotUsed, SuspiciousPredeclaredInstanceAccess, ConstantNotUsed, AssignedByValParameter
'@Folder "Collections"
Option Explicit

'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericHashMap
'*
'* Purpose:
'*
'*
'*======================================================================================================================

Implements IGeneric
Implements IGenericCollection
Implements IGenericSequence

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

Private Const NOT_SPECIFIED As Long = -1 'Dont change!
Private Const LOAD_FACTOR As Double = 0.75
Private Const DEFAULT_CAPACITY As Long = 8

'Linear congruential generator
Private Const INCREMENT As Long = 1
Private Const MULTIPLIER As Long = 5

'Bit Spreading
Private Const BIT_SPREAD As Long = 2 ^ 16

Private Const LOG2 As Double = 0.693147180559945   'Dont change!

Private Type HashItem
    Index As Long
    Code As Long
    Collision As Long
End Type

Private Type Entry
    HashCode As Long
    Key As Object
    Item As Object
End Type

Private Type Member
    HashProvider As IGenericHashProvider
    Occupancy As Long
    Count As Long
    Buckets() As Long
    Entries() As Entry
    Size As Long
    Mask As Long
    Capacity As Long
End Type
Private This As Member

Public Property Get Instance() As IGeneric: Set Instance = Me: End Property
Public Property Get Elements() As IGenericCollection: Set Elements = Me: End Property
Public Property Get HashProvider() As IGenericHashProvider: Set HashProvider = This.HashProvider: End Property
Public Property Get Capacity() As Long: Capacity = This.Capacity: End Property
Public Property Get MeanSearch() As Double: MeanSearch = This.Occupancy / This.Count: End Property

Public Property Get LoadFactor() As Double

    If (This.Capacity > x0000) Then _
        LoadFactor = VBA.Round(This.Count / This.Capacity, 2)

End Property

Public Property Set ElementAt(ByVal Index As Long, ByVal Element As Object)
   
    If (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Public Property Set ElementAt").IndexOutofRange("(Index >= This.Count)")
    
    Set This.Entries(Index).Item = Element

End Property

Public Property Get ElementAt(ByVal Index As Long) As Object
    
    If (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Public Property Get ElementAt").IndexOutofRange("(Index >= This.Count)")
    
    Set ElementAt = This.Entries(Index).Item

End Property

Public Property Get GetAt(ByVal Index As Long) As GenericPair
    
    If (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Public Property Get GetAt").IndexOutofRange("(Index >= This.Count)")
    
    Set GetAt = GenericPair.Build(This.Entries(Index).Key, This.Entries(Index).Item)

End Property

Public Property Get GetKeyAt(ByVal Index As Long) As Object
    
    If (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Public Property Get GetKeyAt").IndexOutofRange("(Index >= This.Count)")
    
    Set GetKeyAt = This.Entries(Index).Key
    
End Property

Public Property Get GetItemAt(ByVal Index As Long) As Object
    
    If (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Public Property Get GetItemAt").IndexOutofRange("(Index >= This.Count)")
    
    Set GetItemAt = This.Entries(Index).Item
    
End Property

Public Property Set Item(ByVal Key As Object, ByVal Element As Object): Call Me.TryAdd(Key, Element): End Property
Public Property Get Item(ByVal Key As Object) As Object

    Dim i As Long
    
    If (This.HashProvider Is Nothing) Then
        i = FindIndex(Key)
    Else
        i = FindIndexByProvider(Key, This.HashProvider)
    End If
    
    If (i < x0000) Then _
        Call GenericError.Build(Me, "Public Property Get Item").NoSuchElement
        
    Set Item = This.Entries(i).Item
    
End Property

Public Function Build(Optional ByVal Capacity As Long, Optional ByVal HashProvider As IGenericHashProvider) As GenericHashMap
    
    Dim newMap As Member
    
    If Not (LOAD_FACTOR >= 0.1 And LOAD_FACTOR <= 1#) Then _
       Call GenericError.Build(Me, "Build").ArgOutOfRange("Not (LoadFactor >= 0.1 And LoadFactor <= 1#")
    
    If (Capacity < x0000) Then _
       Call GenericError.Build(Me, "Build").ArgOutOfRange("Capacity < 0")

    Set newMap.HashProvider = HashProvider
    
    If (Capacity > x0000) Then _
        Call Malloc(newMap, VBA.Int(Capacity / LOAD_FACTOR) + 1)
                                 
    Set Build = New GenericHashMap
    Call System.Inject(Build, VarPtr(newMap), LenB(newMap))
    
End Function

Public Function GetKeys() As GenericArray

    Dim i As Long
    Dim Index As Long
    
    Set GetKeys = GenericArray 'Avoid null
    
    If (This.Count = x0000) Then _
        Exit Function
        
    Set GetKeys = GenericArray.Build(This.Count)
    Index = GenericArray.LowerBound
    
    With GetKeys
        For i = 0 To This.Count - 1: Set .ElementAt(Index + i) = This.Entries(i).Key: Next
    End With
   
End Function

Public Function GetItems() As GenericArray

    Dim i As Long
    Dim Index As Long
    
    Set GetItems = GenericArray 'Avoid null
    
    If (This.Count = x0000) Then _
        Exit Function
        
    Set GetItems = GenericArray.Build(This.Count)
    Index = GenericArray.LowerBound
    
    With GetItems
        For i = 0 To This.Count - 1: Set .ElementAt(Index + i) = This.Entries(i).Item: Next
    End With
    
End Function

Public Function AddAll(ByVal Keys As IGenericSequence) As GenericHashMap

    Dim Key As Object
  
    If (Keys Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function AddAll").IsNothing("Keys")

    With Keys.Iterator
        Do While .HasNext(Key): Call Me.TryAdd(Key, Nothing): Loop
    End With

    Set AddAll = Me
    
End Function

Public Function TryAdd(ByVal Key As Object, ByVal Item As Object) As Boolean

    Dim Entry As Entry
    Dim Hash As HashItem
    Dim i As Long, Slot As Long
    Dim FoundUnusedSlot As Boolean, UnusedSlot As Long
    
    Dim Casted As IGenericValue
    
    If (This.Count = This.Capacity) Then _
        Call Expand
    
    If (This.HashProvider Is Nothing) Then
    
        Set Casted = Key 'Try cast
        
        Hash.Code = Casted.HashCode
        Hash.Code = (Hash.Code Xor (Hash.Code \ BIT_SPREAD))
        Hash.Index = (Hash.Code And This.Mask)
            
        Do
        
            i = This.Buckets(Hash.Index)
            If (i = x0000) Then _
                Exit Do
                
            If Not (i = x80000000) Then
                Slot = (i - 1 And x7FFFFFFF)
                If (This.Entries(Slot).HashCode = Hash.Code) Then _
                    If Casted.Equals(This.Entries(Slot).Key) Then _
                        Set This.Entries(Slot).Item = Item: Exit Function
                
                If Not FoundUnusedSlot Then _
                    This.Buckets(Hash.Index) = (i Or x80000000)
                   
            ElseIf Not FoundUnusedSlot Then
                UnusedSlot = Hash.Index
                FoundUnusedSlot = True
            End If
            
            Hash.Index = (Hash.Index * MULTIPLIER + INCREMENT) And This.Mask
        Loop While (InterlockedIncrement(Hash.Collision) < This.Size)
        
    Else
        
        Hash.Code = This.HashProvider.HashOf(Key)
        Hash.Code = (Hash.Code Xor (Hash.Code \ BIT_SPREAD))
        Hash.Index = (Hash.Code And This.Mask)
        
        Do
        
            i = This.Buckets(Hash.Index)
            If (i = x0000) Then _
                Exit Do
                
            If Not (i = x80000000) Then
                Slot = (i - 1 And x7FFFFFFF)
                If (This.Entries(Slot).HashCode = Hash.Code) Then _
                    If This.HashProvider.Equality(This.Entries(Slot).Key, Key) Then _
                        Set This.Entries(Slot).Item = Item: Exit Function
                
                If Not FoundUnusedSlot Then _
                    This.Buckets(Hash.Index) = (i Or x80000000)
               
            ElseIf Not FoundUnusedSlot Then
                UnusedSlot = Hash.Index
                FoundUnusedSlot = True
            End If
            
            Hash.Index = (Hash.Index * MULTIPLIER + INCREMENT) And This.Mask
        Loop While (InterlockedIncrement(Hash.Collision) < This.Size)
        
    End If
    
    Entry.HashCode = Hash.Code
    Set Entry.Key = Key
    Set Entry.Item = Item
    
    This.Entries(This.Count) = Entry
    This.Count = This.Count + 1
    
    If Not FoundUnusedSlot Then
        This.Buckets(Hash.Index) = This.Count
    Else
        This.Buckets(UnusedSlot) = (This.Count Or x80000000)
    End If
    
    TryAdd = True
    
End Function

Public Function TryGet(ByVal Key As Object, ByRef Item As Object) As Boolean
    
    Dim i As Long
    
    If (This.Count = x0000) Then _
        Exit Function
    
    If (This.HashProvider Is Nothing) Then
        i = FindIndex(Key)
    Else
        i = FindIndexByProvider(Key, This.HashProvider)
    End If
    
    If (i < x0000) Then _
        Exit Function
        
    Set Item = This.Entries(i).Item
    TryGet = True
  
End Function

Public Function TryTake(ByVal Key As Object, ByRef Item As Object) As Boolean: TryTake = Me.TryRemove(Key, Item): End Function

Public Function TryRemove(ByVal Key As Object, Optional ByRef Item As Object) As Boolean
       
    Dim Hash As HashItem
    Dim i As Long, Slot As Long, Previous As Long
    Dim Entry As Entry
    
    Dim Casted As IGenericValue
    
    If (This.Count = x0000) Then _
        Exit Function
        
    If (This.HashProvider Is Nothing) Then
        
        Set Casted = Key
        
        Hash.Code = Casted.HashCode
        Hash.Code = (Hash.Code Xor (Hash.Code \ BIT_SPREAD))
        Hash.Index = (Hash.Code And This.Mask)
      
        Do
        
            i = This.Buckets(Hash.Index)
            
            If (i = x0000) Then _
                Exit Function
                
            If Not (i = x80000000) Then
                Slot = (i - 1 And x7FFFFFFF)
                If (This.Entries(Slot).HashCode = Hash.Code) Then _
                    If Casted.Equals(This.Entries(Slot).Key) Then _
                        Exit Do
            End If
            
            If ((i And x80000000) = x0000) Then _
                Exit Function
            
            Previous = Hash.Index
            Hash.Index = (Hash.Index * MULTIPLIER + INCREMENT) And This.Mask
        Loop While (InterlockedIncrement(Hash.Collision) < This.Size)
        
    Else
    
        Hash.Code = This.HashProvider.HashOf(Key)
        Hash.Code = (Hash.Code Xor (Hash.Code \ BIT_SPREAD))
        Hash.Index = (Hash.Code And This.Mask)
        
        Do
        
            i = This.Buckets(Hash.Index)
            
            If (i = x0000) Then _
                Exit Function
                
            If Not (i = x80000000) Then
                Slot = (i - 1 And x7FFFFFFF)
                If (This.Entries(Slot).HashCode = Hash.Code) Then _
                    If This.HashProvider.Equality(This.Entries(Slot).Key, Key) Then _
                        Exit Do
            End If
            
            If ((i And x80000000) = x0000) Then _
                Exit Function
            
            Previous = Hash.Index
            Hash.Index = (Hash.Index * MULTIPLIER + INCREMENT) And This.Mask
        Loop While (InterlockedIncrement(Hash.Collision) < This.Size)
    
    End If
    
    If (i > x0000) Then ' End of chain/ not marked
    
        This.Buckets(Hash.Index) = x0000
        'Reduce chain
        If (Hash.Collision > x0000) Then _
            This.Buckets(Previous) = (This.Buckets(Previous) And x7FFFFFFF)
    Else
        This.Buckets(Hash.Index) = x80000000 'Mark as part of a chain
    End If
    
    'Save item
    Set Item = This.Entries(Slot).Item
    
    'Shift bucket index / avoid gaps
    If Not (Slot + 1 = This.Count) Then

        Entry = This.Entries(This.Count - 1)
        
        Hash.Code = Entry.HashCode
        Hash.Index = (Hash.Code And This.Mask)

        i = This.Buckets(Hash.Index)

        Do While Not ((i And x7FFFFFFF) = This.Count)
            Hash.Index = (Hash.Index * MULTIPLIER + INCREMENT) And This.Mask
            i = This.Buckets(Hash.Index)
        Loop

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

Public Function Pop() As GenericPair
 
    If (This.Count = x0000) Then _
        Exit Function

    Set Pop = Me.TakeAt(This.Count - 1)
     
End Function

Public Function TakeAt(ByVal Index As Long) As GenericPair

    If (Index >= This.Count) Then _
        Call GenericError.Build(Me, "Public Sub TakeAt").IndexOutofRange("(Index >= This.Count)")
    
    Set TakeAt = GenericPair.Build(This.Entries(Index).Key, This.Entries(Index).Item)
    Call RemoveEntry(Index)
    
End Function

Public Function Contains(ByVal Key As Object) As Boolean:  Contains = Me.IndexOf(Key) >= x0000: End Function

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
                    If (FindIndexByProvider(Candidate, This.HashProvider) < 0) Then _
                        Exit Function
                Loop
                
            End If
        End With
    End If
    
    ContainsAll = True

End Function

Public Function IndexOf(ByVal Key As Object) As Long
   
    If (This.Count > x0000) Then
        If (This.HashProvider Is Nothing) Then
            IndexOf = FindIndex(Key)
        Else
            IndexOf = FindIndexByProvider(Key, This.HashProvider)
        End If
        Exit Function
    End If
    
    IndexOf = -1
    
End Function

Private Sub RemoveEntry(ByVal Index As Long)
    
    Dim Entry As Entry
    Dim Hash As HashItem
    Dim i As Long, Bucket As Long, Previous As Long
    
    Hash.Code = This.Entries(Index).HashCode
    Hash.Index = (Hash.Code And This.Mask)
    
    Bucket = Index + 1
    Previous = Bucket
    
    i = This.Buckets(Hash.Index)
   
    Do While Not ((i And x7FFFFFFF) = Bucket)
        Previous = Hash.Index
        Hash.Index = (Hash.Index * MULTIPLIER + INCREMENT) And This.Mask
        i = This.Buckets(Hash.Index)
    Loop
    
    If (i > x0000) Then
        If Not (Previous = Bucket) Then _
            This.Buckets(Previous) = This.Buckets(Previous) And x7FFFFFFF
        This.Buckets(Hash.Index) = x0000
    Else
        This.Buckets(Hash.Index) = x80000000
    End If
    
    If Not (Bucket = This.Count) Then
        
        Entry = This.Entries(This.Count - 1)
        
        Hash.Code = Entry.HashCode
        Hash.Index = (Hash.Code And This.Mask)
        
        i = This.Buckets(Hash.Index)
             
        Do While Not ((i And x7FFFFFFF) = This.Count)
            Hash.Index = (Hash.Index * MULTIPLIER + INCREMENT) And This.Mask
            i = This.Buckets(Hash.Index)
        Loop

        If (i < x0000) Then _
            Bucket = (Bucket Or x80000000)
        
        This.Buckets(Hash.Index) = Bucket
        This.Entries(Index) = Entry
    
    End If

    This.Count = This.Count - 1
    This.Entries(This.Count) = EmptyEntry
    
End Sub

Public Function RemoveAll(ByVal Keys As IGenericSequence) As Long

    Dim Candidate As Object
    
    If (Keys Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function RemoveAll").IsNothing("(Keys Is Nothing)")
    
    If (This.Count = x0000) Then _
        Exit Function
        
    If (Keys Is Me) Then
    
        RemoveAll = This.Count
        Call Me.Elements.Clear
        
    Else

        With Keys.Iterator
            Do While .HasNext(Candidate)
                If (Me.TryRemove(Candidate)) Then _
                    RemoveAll = RemoveAll + 1
            Loop
        End With
        
    End If
    
End Function

Public Function Slice(ByVal BeginAt As Long, Optional ByVal Count As Long = NOT_SPECIFIED) As GenericHashMap
    
    Dim SubMap As GenericHashMap
    Dim i As Long
    
    If (BeginAt < x0000) Then _
        Call GenericError.Build(Me, "Public Function Slice").ArgOutOfRange("(BeginAt < 0)")
        
    If (Count = NOT_SPECIFIED) Then _
        Count = This.Count - BeginAt
   
    If (Count <= x0000) Then _
        Call GenericError.Build(Me, "Public Function Slice").ArgOutOfRange("(Count <= 0)")
        
    If (This.Count - BeginAt < Count) Then _
        Call GenericError.Build(Me, "Public Function Slice").InvalidInput("(This.Count - BeginAt < Count)")
   
    Set SubMap = GenericHashMap.Build(Capacity:=Count, HashProvider:=This.HashProvider)
    
    With SubMap
        For i = BeginAt To BeginAt + Count - 1: Call .TryAdd(This.Entries(i).Key, This.Entries(i).Item): Next
    End With
    
    Set Slice = SubMap
    
End Function

Public Function LastEntry() As GenericPair
    
    If (This.Count = x0000) Then _
        Exit Function
    
    Set LastEntry = GenericPair.Build(This.Entries(This.Count - 1).Key, This.Entries(This.Count - 1).Item)
    
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
    
    '(MULTIPLIER < Table.Length)!
    If (MULTIPLIER > Size) Then _
        Size = MULTIPLIER + 1

    Size = GetPowerOfTwo(Size)
    
    Struct.Occupancy = x0000
    Struct.Capacity = VBA.Int(Size * LOAD_FACTOR)
    
    If Not (Struct.Capacity < Size) Then _
        Call GenericError.Build(Me, "Private Function Malloc").InvalidInput("Based on the current algorithm, Capacity must be less than Size.")
        
    Struct.Mask = Size - 1
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

Private Function GetPowerOfTwo(ByVal Number As Long) As Long
    
    If IsPowerOfTwo(Number) Then
        GetPowerOfTwo = Number
    Else 'You can calculate base-n logarithms for any number x by dividing the natural logarithm of x by the natural logarithm of n as follows: LOGn(x) = Log(x) / Log(n)
        GetPowerOfTwo = CLng(2 ^ VBA.Abs(VBA.Int(-VBA.Log(Number) / LOG2)))
    End If
    
End Function

Private Function IsPowerOfTwo(ByVal Number As Long) As Boolean: IsPowerOfTwo = ((Number And (Number - 1)) = 0): End Function

'Private Function PowerOfTwo(ByVal Number As Long) As Long
'
'    Dim pos As Long
'
'    Do While (Number > 0)
'        pos = pos + 1
'        Number = Number \ 2
'    Loop
'
'    getPowerOfTwo = 2 ^ pos
'
'End Function
'@Ignore EmptyMethod
Private Function EmptyEntry() As Entry: End Function

Public Function TrimExcess() As GenericHashMap
    
    Dim Size As Long
    
    If (This.Capacity > x0000) Then
            
        Size = GetPowerOfTwo(VBA.Int(This.Count / LOAD_FACTOR) + 1)
        
        If (Size < This.Size) Then _
            Call Resize(Size)
                
    End If
    
    Set TrimExcess = Me
    
End Function

Public Function EnsureCapacity(ByVal Min As Long) As GenericHashMap
    
    Dim Size As Long
    
    If (Me Is GenericHashMap) Then _
        Call GenericError.Build(Me, "Public Function EnsureCapacity").InvalidOperation("(Me Is GenericHashMap)")
        
    Size = VBA.Int(Min / LOAD_FACTOR) + 1
    
    If (Size > This.Capacity) Then _
        Call Resize(Size)
    
    Set EnsureCapacity = Me
    
End Function

Private Sub Expand()
    
    Dim Size As Long
    
    If (Me Is GenericHashMap) Then _
        Call GenericError.Build(Me, "Private Sub Expand").InvalidOperation("(Me Is GenericHashMap)")
        
    If (This.Size > x0000) Then
        Size = This.Size * 2
    Else
        Size = VBA.Int(DEFAULT_CAPACITY / LOAD_FACTOR) + 1
    End If
    
    Call Resize(Size)
    
End Sub

Public Function Rehash(Optional ByVal ForceRehashing As Boolean = False) As GenericHashMap
 
    Dim i As Long
   
    Set Rehash = Me
    
    If (This.Capacity = x0000) Then _
        Exit Function
    
    If ForceRehashing Then
        If Not (This.HashProvider Is Nothing) Then
            With This.HashProvider
                For i = 0 To This.Count - 1: This.Entries(i).HashCode = .HashOf(This.Entries(i).Key): Next
            End With
        End If
    End If
    
    ReDim This.Buckets(This.Size - 1)
    Call LoadBuckets
         
End Function

Private Sub Resize(ByVal Size As Long)
   
    'Clear cache and allocate new
    Call Malloc(This, Size)
    
    If (This.Count = x0000) Then _
        Exit Sub
        
    Call LoadBuckets
                                                                                            
End Sub

Private Sub LoadBuckets()

    Dim Hash As HashItem
    Dim i As Long, Slot As Long

    This.Occupancy = x0000

    For i = 1 To This.Count
    
        Hash.Code = This.Entries(i - 1).HashCode
        Hash.Index = (Hash.Code And This.Mask)
        Slot = This.Buckets(Hash.Index)
    
        Do While Not (Slot = x0000)
            This.Buckets(Hash.Index) = (Slot Or x80000000)
            Hash.Index = (Hash.Index * MULTIPLIER + INCREMENT) And This.Mask
            Slot = This.Buckets(Hash.Index)
        Loop

        This.Buckets(Hash.Index) = i
        
    Next

End Sub

'Try cast
Private Function FindIndex(ByVal Key As IGenericValue) As Long

    Dim Hash As HashItem
    Dim i As Long
    
    Hash.Code = Key.HashCode
    Hash.Code = (Hash.Code Xor (Hash.Code \ BIT_SPREAD))
    Hash.Index = (Hash.Code And This.Mask)
     
    Do
        
        i = This.Buckets(Hash.Index)
        
        If (i = x0000) Then _
            Exit Do
      
        If Not (i = x80000000) Then
            FindIndex = (i - 1 And x7FFFFFFF)
            
            If (This.Entries(FindIndex).HashCode = Hash.Code) Then
                
                If Key Is This.Entries(FindIndex).Key Then _
                    Exit Function
                
                If (Key.Equals(This.Entries(FindIndex).Key)) Then
                    Set This.Entries(FindIndex).Key = Key
                    Exit Function
                End If
            
            End If
                    
'            (Bucket >= 0)
            If ((i And x80000000) = x0000) Then _
                Exit Do
           
        End If
        
        Hash.Index = (Hash.Index * MULTIPLIER + INCREMENT) And This.Mask
    Loop While (InterlockedIncrement(Hash.Collision) < This.Size)
    
    FindIndex = -1
    
End Function

Private Function FindIndexByProvider(ByVal Key As Object, ByVal HashProvider As IGenericHashProvider) As Long

    Dim Hash As HashItem
    Dim i As Long
    
    Hash.Code = HashProvider.HashOf(Key)
    Hash.Code = (Hash.Code Xor (Hash.Code \ BIT_SPREAD))
    Hash.Index = (Hash.Code And This.Mask)
     
    Do
        
        i = This.Buckets(Hash.Index)
        
        If (i = x0000) Then _
            Exit Do
            
        If Not (i = x80000000) Then
            FindIndexByProvider = (i - 1 And x7FFFFFFF)
            
            If (This.Entries(FindIndexByProvider).HashCode = Hash.Code) Then _
                If (HashProvider.Equality(This.Entries(FindIndexByProvider).Key, Key)) Then _
                    Exit Function
            '(Bucket >= 0)
            If ((i And x80000000) = x0000) Then _
                Exit Do
                
        End If
        
        Hash.Index = (Hash.Index * MULTIPLIER + INCREMENT) And This.Mask
    Loop While (InterlockedIncrement(Hash.Collision) < This.Size)

    FindIndexByProvider = -1
    
End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericHashMap: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = GenericHashMap: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(This): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(This): End Property
Private Property Get IGeneric_HashCode() As Long:  IGeneric_HashCode = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericHashMap):  End Property

Private Function IGeneric_Equals(ByVal Other As Object) As Boolean

    If (Other Is Nothing) Then _
        Exit Function
    
    If TypeOf Other Is GenericHashMap Then _
        IGeneric_Equals = Me Is Other
    
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
    
    Set IGenericCollection_GetAt = GenericPair.Build(This.Entries(Index).Key, This.Entries(Index).Item)

End Function

Private Function IGenericCollection_Contains(ByVal Element As Object) As Boolean

    If (This.Count = x0000) Then _
        Exit Function
    
    If (This.HashProvider Is Nothing) Then
        IGenericCollection_Contains = FindIndex(Element) >= 0
    Else
        IGenericCollection_Contains = FindIndexByProvider(Element, This.HashProvider) >= 0
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
        For i = 0 To This.Count - 1: Set .ElementAt(Index + i) = GenericPair.Build(This.Entries(i).Key, This.Entries(i).Item): Next
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
        .Mask = This.Mask
        .Occupancy = This.Occupancy
        .Buckets = This.Buckets
        .Entries = This.Entries
    End With
    
    Set IGenericCollection_Copy = New GenericHashMap
    Call System.Inject(IGenericCollection_Copy, VarPtr(Clone), LenB(Clone))
    
End Function

Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Range: End Function

'IGenericSequence_Iterator
Private Function IGenericSequence_Iterator() As IGenericIterator: Set IGenericSequence_Iterator = Me.Range: End Function


