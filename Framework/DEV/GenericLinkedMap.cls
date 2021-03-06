VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericLinkedMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericLinkedMap
'*
'* Purpose: A GenericLinkedMap represents a data structure that can store objects as key value pairs.
'*          It is dynamic in nature means the size of the map is growing according to the need.
'*          You can search for a value in an instance of GenericLinkedMap class using the corresponding key.
'*          It holds both keys and values (A key cannot be null, but a value can be) in a
'*          pseudo random order determined by hashes calculated from keys.
'*          It is also sparse, with unoccupied holes in a pre-allocated array.
'*          A GenericLinkedMap object consists of buckets that contain the elements of the collection.
'*          A bucket is a virtual subgroup of elements within the map, which makes searching and retrieving easier and faster than in most collections.
'*          Each bucket is associated with a hash code, which is generated using a hash function and is based on the key of the element.
'*          Key objects must be immutable as long as they are used as keys in the map.
'*          When an element is added to the map, the element is placed into a bucket based on the hash code of the key.
'*          Subsequent lookups of the key use the hash code of the key to search in only one particular bucket,
'*          thus substantially reducing the number of key comparisons required to find an element.
'*          The load factor of a map determines the maximum ratio of elements to buckets.
'*          Smaller load factors cause faster average lookup times at the cost of increased memory consumption.
'*          The default load factor of 0.72 generally provides the best balance between speed and size.
'*          A different load factor can also be specified when the map is created.
'*          As elements are added to a map, the actual load factor of the map increases.
'*          When the actual load factor reaches the specified load factor, the number of buckets in the map is automatically
'*          increased to the smallest prime number that is larger than twice the current number of Hashtable buckets.
'*          If a specified key already exists in the map, setting the Item() property overwrites the old value.
'*          In contrast, the Add method throws an exception if an element with the specified key already exists.
'*
'* Notes:   Key objects must be immutable as long as they are used as keys in the Hashtable.
'*          A key cannot be null, but a value can be.
'*          Each key object in the map must provide its own hash function.
'*          A GenericLinkedMap does maintain the order of stored values.
'*          This map uses double hashing. There are hashsize buckets in the map, and ech bucket can contain 0 or 1 element.
'*
'*======================================================================================================================

'@Exposed
'@PredeclaredId
'@Folder "<T>Enumerable"
Option Explicit

Implements IGeneric
Implements IGenericMap
Implements IGenericCollection
Implements IGenericIterator

#If Win64 Then
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

Private Const LOAD_FACTOR As Double = 0.7
Private Const MAXPRIME As Long = 2 ^ 31 - 1
Private Const HASHPRIME As Long = 101
Private Const DEFAULT_CAPACITY As Long = 32

'Bucketstate
Private Const IS_EMPTY As Byte = 0
Private Const HAS_ENTRY As Byte = 2 ^ 0
Private Const HAS_COLLISION As Byte = 2 ^ 1

Private Type HashItem
    Incr As Long
    Size As Long
    Index As Long
    Code As Long
End Type
'
Private Type Iterator
    HasLoaded As Boolean
    current As IGeneric
    Head As GenericBucket
    Bucket As GenericBucket
End Type
'
Private Type Member
    Comparer As IGenericComparer
    Capacity As Long
    Occupancy As Long
    Count As Long
    Size As Long
    LoadSize As Long
    Buckets As GenericArray
    Head As GenericBucket
    Iterator As Iterator
End Type
Private myclass As Member

Public Property Get Elements() As IGenericCollection: Set Elements = Me: End Property
Public Property Get Comparer() As IGenericComparer: Set Comparer = myclass.Comparer: End Property
Public Property Set Item(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value, True): End Property
Public Property Get Item(ByVal Key As IGeneric) As IGeneric
    
    Dim Bucket As GenericBucket: Set Bucket = FindBucket(Key)
    
    If (Bucket Is Nothing) Then _
        Call System.Throw(Me, "Public Property Get Item").NoSuchElement
    
    Set Item = Bucket.Value

End Property

Public Function Build(Optional ByVal Capacity As Long, Optional ByVal Comparer As IGenericComparer, Optional Head As GenericBucket) As GenericLinkedMap
    
    Dim newMap As Member
    Dim RawSize As Long
    
    If Not (LOAD_FACTOR >= 0.1 And LOAD_FACTOR <= 1#) Then _
       Call System.Throw(Me, "Build").ArgOutOfRange("Not (LoadFactor >= 0.1 And LoadFactor <= 1#")
    
    If (Capacity < 0) Then _
       Call System.Throw(Me, "Build").ArgOutOfRange("Capacity < 0")
    
    If Comparer Is Nothing Then _
        Set Comparer = IGenericComparer
    Set newMap.Comparer = Comparer
    
    'Avoid null
    Set newMap.Buckets = GenericArray
    
    Set newMap.Head = New GenericBucket
    Set newMap.Head.Before = newMap.Head
    Set newMap.Head.After = newMap.Head
    Set Head = newMap.Head
    
    If (Capacity > 0) Then _
        Call Initialize(newMap, (Capacity / LOAD_FACTOR))
                                 
    Set Build = System.CreateInstance(New GenericLinkedMap, VarPtr(newMap), LenB(newMap))
    
End Function

Public Function ContainsKey(ByVal Key As IGeneric) As Boolean: ContainsKey = FindEntry(Key) > -1: End Function 'TryGetValue(Key, Nothing): End Function
'Use equality
Public Function ContainsValue(ByVal Value As IGeneric) As Boolean

    Dim Bucket As GenericBucket
    Set Bucket = myclass.Head.After
    
    If (Value Is Nothing) = False Then
        Do While (Bucket Is myclass.Head = False)
            If Value.Equals(Bucket.Value) Then _
                ContainsValue = True: Exit Function
            Set Bucket = Bucket.After
        Loop
    Else
        Do While (Bucket Is myclass.Head = False)
            If (Bucket.Value Is Nothing) Then _
                ContainsValue = True: Exit Function
            Set Bucket = Bucket.After
        Loop
    End If

End Function

Public Function GetKeys() As GenericArray

    Dim Keys As GenericArray
    Dim Index As Long
    Dim Bucket As GenericBucket
    Set Bucket = myclass.Head.After
    
    Set Keys = GenericArray.Build(Me.Elements.Count)
    
    With Keys
        Index = .LowerBound - 1
        Do While (Bucket Is myclass.Head = False)
            Set .ElementAt(InterlockedIncrement(Index)) = Bucket.Key
            Set Bucket = Bucket.After
        Loop
    End With
    
    Set GetKeys = Keys

End Function

Public Function GetValues() As GenericArray

    Dim Values As GenericArray
    Dim Index As Long
    Dim Bucket As GenericBucket
    Set Bucket = myclass.Head.After
    
    Set Values = GenericArray.Build(Me.Elements.Count)
    
    With Values
        Index = .LowerBound - 1
        Do While (Bucket Is myclass.Head = False)
            Set .ElementAt(InterlockedIncrement(Index)) = Bucket.Value
            Set Bucket = Bucket.After
        Loop
    End With
    
    Set GetValues = Values
    
End Function

Public Function TryGetValue(ByVal Key As IGeneric, ByRef Value As IGeneric) As Boolean
             
    Dim Bucket As GenericBucket: Set Bucket = FindBucket(Key)
    
    If Not (Bucket Is Nothing) Then _
        Set Value = Bucket.Value: TryGetValue = True
  
End Function

Public Function Add( _
                    ByVal Key As IGeneric, _
                    ByVal Value As IGeneric, _
                    Optional ByVal Overwrite As Boolean) As Boolean

    Dim Try As Long
    Dim Hash As HashItem
    Dim Bucket As GenericBucket, State As Byte
    Dim EmptyBucket As GenericBucket
   
    If (Key Is Nothing) Then _
        Call System.Throw(Me, "Public Function Add").IsNothing("(Key Is Nothing)")
        
    With myclass
    
        If (.Count >= .LoadSize) Then
                                                                                                                      
            If .Size > 0 Then
                Call Rehash(ExpandPrime(.Size), False)
            Else
                Call Initialize(myclass, DEFAULT_CAPACITY)
            End If
        
        ElseIf (.Occupancy > .LoadSize) Then
            Call Rehash(getPrime(.Size + 1), False)
        
        End If
        
        Call LoadHash(Hash, Key.HashValue, .Size) 'TODO Should we?! myclass.Comparer.GetHashCode(Key)
    
        Do
            
            State = IS_EMPTY '!
            Set Bucket = .Buckets.ElementAt(Hash.Index)
            
            If Bucket Is Nothing = False Then
                State = Bucket.State
                                                    
                If (EmptyBucket Is Nothing) Then _
                    If Not (State And HAS_ENTRY) = HAS_ENTRY Then _
                        Set EmptyBucket = Bucket
                
            End If
            
            If (State = IS_EMPTY) Then
            
                If EmptyBucket Is Nothing Then
                    Set EmptyBucket = New GenericBucket
                    Set myclass.Buckets.ElementAt(Hash.Index) = EmptyBucket
                End If
                
                With EmptyBucket
                    Set .After = myclass.Head
                    .State = (HAS_ENTRY Or .State) '!
                    .HashCode = Hash.Code
                    Set .Key = Key: Set .Value = Value
                End With
                
                With myclass.Head
                    Set EmptyBucket.Before = .Before
                    Set .Before.After = EmptyBucket
                    Set .Before = EmptyBucket
                End With
                   
                Call InterlockedIncrement(myclass.Count)
                Add = True
                Exit Function
                
            End If
            
            If Bucket.HashCode = Hash.Code Then
                If .Comparer.Equals(Bucket.Key, Key) Then
                    If (Overwrite) Then _
                        Set Bucket.Value = Value
                    Exit Function
                End If
            End If
            
            If (EmptyBucket Is Nothing) Then _
                If (Not (State And HAS_COLLISION) = HAS_COLLISION) Then _
                    Bucket.State = (HAS_COLLISION Or State): Call InterlockedIncrement(myclass.Occupancy)

            Hash.Index = Modulo(Hash.Index + Hash.Incr, Hash.Size)

        Loop While (InterlockedIncrement(Try) < Hash.Size)
    
    End With
    
    Call System.Throw(Me, "Public Function Add").InvalidOperation("Can't add :: [" & Key & ";" & Value & "]")
    
End Function

Public Sub AddAll(ByVal Map As IGenericMap, Optional ByVal Overwrite As Boolean)

    Dim Pair As GenericPair
  
    If (Map Is Nothing) Then _
        Call System.Throw(Me, "Public Function AddAll").IsNothing("(Map Is Nothing)")
    
    If (Map Is Me) Then _
        Exit Sub
            
    If (Map.Elements.Count > 0) Then
        With Map.Elements.Iterator
            Do While .HasNext(Pair): Call Me.Add(Pair.Key, Pair.Value, Overwrite:=Overwrite): Loop
        End With
    End If

End Sub

Public Function Remove(ByVal Key As IGeneric) As Boolean
    
    Dim Bucket As GenericBucket: Set Bucket = FindBucket(Key)
    
    If Not (Bucket Is Nothing) Then
        With Bucket 'remain as empty bucket
            .State = (.State And Not HAS_ENTRY)
            .HashCode = -1
            Set .Key = Nothing: Set .Value = Nothing
        
            Set .After.Before = .Before
            Set .Before.After = .After
            Set .Before = Nothing
            Set .After = Nothing
        End With
        
        Call InterlockedDecrement(myclass.Count)
        Remove = True
    End If
    
End Function

Public Function RemoveAll(ByVal Collection As IGenericCollection) As Long

    Dim Candidate As IGeneric
    Dim Count As Long
    
    If (Collection Is Nothing) Then _
        Call System.Throw(Me, "Public Function RemoveAll").IsNothing("(Collection Is Nothing)")
    
    If (Collection.Count > 0) Then
        
        If (Collection Is Me) Then
            Call Me.Elements.Clear
            
        Else

            With Collection.Iterator
                Do While .HasNext(Candidate)
                    If Me.Remove(Candidate) Then _
                        Count = Count + 1
                Loop
            End With
            
        End If
    End If

    RemoveAll = Count
    
End Function

Private Sub Initialize(ByRef Struct As Member, ByVal Capacity As Long)
    
    If (Capacity > MAXPRIME) Then _
        Call System.Throw(Me, "private Sub Initialize").Overflow("Capacity > MAXPRIME")
    
    Struct.Size = getPrime(Capacity)
    Struct.LoadSize = Int(LOAD_FACTOR * Struct.Size)
                                      
    If Not (Struct.LoadSize < Struct.Size) Then _
        Call System.Throw(Me, "Build").InvalidInput("Based on the current algorithm, loadSize must be less than Size.")
    
    Set Struct.Buckets = GenericArray.Build(Struct.Size)
                                                                                          
End Sub

Private Static Function getPrime(Optional ByVal min As Long = 0) As Long
    
    Dim p As Long
    Dim i As Long
    Dim Primes() As Long
    
    If min < 0 Then _
        Call System.Throw(Me, "getPrime").InvalidInput("Min < 0")
    
    If i = 0 Then
        Dim Prime As Variant
        ReDim Primes(71)
            
        For Each Prime In Array(3, 7, 11, 17, 23, 29, 37, 47, 59, 71, 89, 107, 131, 163, 197, 239, 293, 353, 431, 521, 631, 761, 919, _
                            1103, 1327, 1597, 1931, 2333, 2801, 3371, 4049, 4861, 5839, 7013, 8419, 10103, 12143, 14591, _
                            17519, 21023, 25229, 30293, 36353, 43627, 52361, 62851, 75431, 90523, 108631, 130363, 156437, _
                            187751, 225307, 270371, 324449, 389357, 467237, 560689, 672827, 807403, 968897, 1162687, 1395263, _
                            1674319, 2009191, 2411033, 2893249, 3471899, 4166287, 4999559, 5999471, 7199369)
                
            Primes(i) = Prime
            i = i + 1
        Next
    End If
    
    For p = LBound(Primes) To UBound(Primes) Step 2
        If Primes(p) >= min Then _
            getPrime = Primes(p): Exit Function
    Next
    
    For p = min To MAXPRIME
        If IsPrime(p) Then _
            If (((p - 1) Mod HASHPRIME) = 0 = False) Then _
                getPrime = p: Exit Function
    Next
   
    getPrime = min
    
End Function

Private Function IsPrime(ByVal Candidate As Long) As Boolean
    
    Dim Limit As Long: Limit = sqr(Candidate)
    Dim Divisor As Long
    
    If (Candidate And 1) <> 0 Then
        For Divisor = 3 To Limit Step 2
            If (Candidate Mod Divisor) = 0 Then _
                IsPrime = False: Exit Function
        Next
        IsPrime = True: Exit Function
    End If
    
    IsPrime = CBool(Candidate = 2)
    
End Function

Private Function ExpandPrime(ByVal oldSize As Long) As Long
    
    Dim newSize As Long: newSize = oldSize * 2
    
    If newSize < MAXPRIME Then
        ExpandPrime = getPrime(newSize)
    Else
        ExpandPrime = MAXPRIME
    End If
    
End Function

Private Sub LoadHash(ByRef h As HashItem, ByVal HashCode As Long, ByVal Size As Long)
    
    With h
        .Code = HashCode
        .Size = Size
        .Index = Modulo(.Code, .Size)
        .Incr = 1 + Modulo((CDbl(.Code) * HASHPRIME), (.Size - 1))
    End With
    
End Sub

Private Sub Rehash(ByVal Size As Long, ByVal ForceNewHashing As Boolean)

    Dim NewBuckets As GenericArray
    Set NewBuckets = GenericArray.Build(Size)
   
    Dim oldBucket As GenericBucket
    Dim newBucket As GenericBucket
    Dim Hash As HashItem
   
    myclass.Occupancy = 0
    Set oldBucket = myclass.Head.After
        
    Do While oldBucket Is myclass.Head = False
        
        If Not ForceNewHashing Then
            Call LoadHash(Hash, oldBucket.HashCode, Size)
        Else
            Call LoadHash(Hash, oldBucket.Key.HashValue, Size)
        End If
            
        oldBucket.State = HAS_ENTRY
            
        Do
            Set newBucket = NewBuckets.ElementAt(Hash.Index)
                
            If newBucket Is Nothing Then _
                Set NewBuckets.ElementAt(Hash.Index) = oldBucket: Exit Do
                
            If Not (newBucket.State And HAS_COLLISION) = HAS_COLLISION Then _
                newBucket.State = (HAS_COLLISION Or HAS_ENTRY): Call InterlockedIncrement(myclass.Occupancy)
                
            Hash.Index = Modulo(Hash.Index + Hash.Incr, Hash.Size)
        Loop
            
         Set oldBucket = oldBucket.After
            
    Loop
                                                                                                          
    Set myclass.Buckets = NewBuckets
    myclass.LoadSize = Int(LOAD_FACTOR * Size)
    myclass.Size = Size
    
End Sub

Private Function FindBucket(ByVal Key As IGeneric) As GenericBucket

    Dim Try As Long
    Dim Hash As HashItem
    Dim Bucket As GenericBucket
    
    If (Key Is Nothing) Then _
        Call System.Throw(Me, "Private Function FindBucket").IsNothing("(Key Is Nothing)")

    Call LoadHash(Hash, Key.HashValue, myclass.Size) 'TODO Should we?! myclass.Comparer.GetHashCode(Key)
    
    With myclass.Buckets
        Do
            
            Set Bucket = .ElementAt(Hash.Index)
            
            If (Bucket Is Nothing) Then _
                Exit Function
                                        
            With Bucket
                If .HashCode = Hash.Code Then _
                    If myclass.Comparer.Equals(.Key, Key) Then _
                        FindBucket = Bucket: Exit Function
        
                If (Not (.State And HAS_COLLISION) = HAS_COLLISION) Then _
                    Exit Function
            
            End With
            
            If ((InterlockedIncrement(Try)) > Hash.Size) Then _
                Exit Function
    
            Hash.Index = Modulo(Hash.Index + Hash.Incr, Hash.Size)
            
        Loop
    End With
  
End Function

Private Function FindEntry(ByVal Key As IGeneric) As Long

    Dim Try As Long
    Dim Hash As HashItem
    Dim Bucket As GenericBucket
    
    If (Key Is Nothing) Then _
        Call System.Throw(Me, "Private Function FindEntry").IsNothing("(Key Is Nothing)")

    Call LoadHash(Hash, Key.HashValue, myclass.Size) 'TODO Should we?! myclass.Comparer.GetHashCode(Key)
    
    With myclass.Buckets
        Do
            
            Set Bucket = .ElementAt(Hash.Index)
            
            If (Bucket Is Nothing) Then _
                FindEntry = -1: Exit Function
                                        
            With Bucket
    '            If .State = IS_EMPTY Then _
    '                FindEntry = -1: Exit Function
    
                If .HashCode = Hash.Code Then _
                    If myclass.Comparer.Equals(.Key, Key) Then _
                        FindEntry = Hash.Index: Exit Function
        
                If (Not (.State And HAS_COLLISION) = HAS_COLLISION) Then _
                    FindEntry = -1: Exit Function
            
            End With
            
            If ((InterlockedIncrement(Try)) > Hash.Size) Then _
                FindEntry = -1: Exit Function
    
            Hash.Index = Modulo(Hash.Index + Hash.Incr, Hash.Size)
            
        Loop
    End With
    
    FindEntry = -1
   
End Function

Private Sub LoadIterator( _
                            ByRef Iterator As Iterator, _
                            ByVal Head As GenericBucket)
    
    With Iterator
        Set .Head = Head
        Set .Bucket = Head.After
        .HasLoaded = True
    End With
    
End Sub

Private Sub DisposeLinkedBuckets(ByVal Head As GenericBucket)
    
    Dim Bucket As GenericBucket
    Dim oldBucket As GenericBucket

    If Head Is Nothing Then _
        Exit Sub
    
    Set Bucket = Head.After
    
    Do While Bucket Is Head = False
    Set oldBucket = Bucket
        Set Bucket = oldBucket.After
        Call oldBucket.Invalidate
    Loop
    
    Call Head.Invalidate
    
End Sub

Private Sub Class_Terminate()
    Call DisposeLinkedBuckets(myclass.Head)
    myclass = EmptyStructure
End Sub

Private Function Modulo(ByVal A As Double, ByVal m As Double) As Long: Modulo = (A - (Int(A / m) * m)): End Function
Private Function Decrement(ByRef i As Long) As Long: i = (Not -i): Decrement = i: End Function
Private Function Increment(ByRef i As Long) As Long: i = (-(Not i)): Increment = i: End Function
Private Property Get LoadSize() As Long: LoadSize = myclass.LoadSize: End Property
Private Property Get LoadFactor() As Double: LoadFactor = LOAD_FACTOR: End Property

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericLinkedMap: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GenericLinkedMap: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericLinkedMap):  End Property
Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = Me.Elements.Copy: End Function

Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    IGeneric_Equals = (Me Is Other)
End Function

Private Function IGeneric_ToString() As String
    With myclass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "Size: " & .Size & vbNewLine & _
                            "LoadFactor: " & LOAD_FACTOR & vbNewLine & _
                            "LoadSize: " & .LoadSize & vbNewLine & _
                            "Elements: " & .Count & vbNewLine & _
                            "Occupancy: " & .Occupancy
    End With
End Function

'IGenericMap
Private Property Get IGenericMap_Elements() As IGenericCollection: Set IGenericMap_Elements = Me: End Property
Private Property Get IGenericMap_Comparer() As IGenericComparer: Set IGenericMap_Comparer = myclass.Comparer: End Property
Private Property Get IGenericMap_Item(ByVal Key As IGeneric) As IGeneric: Set IGenericMap_Item = Me.Item(Key:=Key): End Property
Private Property Set IGenericMap_Item(ByVal Key As IGeneric, ByVal Item As IGeneric): Set Me.Item(Key:=Key) = Item: End Property

Private Function IGenericMap_TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: IGenericMap_TryGetValue = Me.TryGetValue(Key:=Key, Value:=Value): End Function
Private Function IGenericMap_GetKeys() As GenericArray: Set IGenericMap_GetKeys = Me.GetKeys: End Function
Private Function IGenericMap_GetValues() As GenericArray: Set IGenericMap_GetValues = Me.GetValues: End Function
Private Function IGenericMap_Add(ByVal Key As IGeneric, ByVal Value As IGeneric, Optional ByVal Overwrite As Boolean) As Boolean: IGenericMap_Add = Me.Add(Key:=Key, Value:=Value, Overwrite:=True): End Function
Private Sub IGenericMap_AddAll(ByVal Map As IGenericMap, Optional ByVal Overwrite As Boolean): Call Me.AddAll(Map:=Map, Overwrite:=Overwrite): End Sub
Private Function IGenericMap_ContainsKey(ByVal Key As IGeneric) As Boolean: IGenericMap_ContainsKey = Me.ContainsKey(Key:=Key): End Function
Private Function IGenericMap_ContainsValue(ByVal Value As IGeneric) As Boolean: IGenericMap_ContainsValue = Me.ContainsValue(Value:=Value): End Function
Private Function IGenericMap_Remove(ByVal Key As IGeneric) As Boolean: IGenericMap_Remove = Me.Remove(Key:=Key): End Function
Private Function IGenericMap_RemoveAll(ByVal Collection As IGenericCollection) As Long: IGenericMap_RemoveAll = Me.RemoveAll(Collection:=Collection): End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = myclass.Count: End Property
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean

    Dim Bucket As GenericBucket
    Dim Pair As GenericPair
    
    If Not Item Is Nothing Then
        If Not TypeOf Item Is GenericPair Then _
            Exit Function
        Set Pair = Item
            
        Set Bucket = FindBucket(Key:=Pair.Key)
        
        If Not Bucket Is Nothing Then _
            IGenericCollection_Contains = IGenericComparer.Equals(Bucket.Value, Pair.Value)
    End If
        
End Function

Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long)

    Dim Bucket As GenericBucket
    
    If (GArray Is Nothing) Then _
        Call System.Throw(Me, "Private Sub IGenericCollection_CopyTo").IsNothing("(GArray Is Nothing)")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex >= GArray.Length) Then _
        Call System.Throw(Me, "Private Sub IGenericCollection_CopyTo").IndexOutofRange("(ArrayIndex < GArray.LowerBound) Or (ArrayIndex >= GArray.Length)")
        
    Set Bucket = myclass.Head.After
    ArrayIndex = ArrayIndex - 1
    
    With GArray
        Do While (Bucket Is myclass.Head = False)
            Set .ElementAt(InterlockedIncrement(ArrayIndex)) = Bucket.ToPair
            Set Bucket = Bucket.After
        Loop
    End With
End Sub

Private Sub IGenericCollection_Clear()

    Call DisposeLinkedBuckets(myclass.Head)
    Set myclass.Head.Before = myclass.Head
    Set myclass.Head.After = myclass.Head
    'dont call myclass.Buckets.clear '->Reuse Buckets
    myclass.Count = 0
    myclass.Occupancy = 0
End Sub

Private Function IGenericCollection_ToArray() As GenericArray

    Dim Pairs As GenericArray
    Set Pairs = GenericArray.Build(Me.Elements.Count)
    
    Call Me.Elements.CopyTo(Pairs, Pairs.LowerBound)
    
    Set IGenericCollection_ToArray = Pairs
    
End Function

Private Function IGenericCollection_Copy() As IGenericCollection
    
    Dim c As Member
    Dim oldBucket As GenericBucket
    Dim newBucket As GenericBucket
    Dim Hash As HashItem
    
    With c
        Set .Comparer = myclass.Comparer
        .Size = myclass.Size
        .Occupancy = 0
        .LoadSize = myclass.LoadSize
        .Count = myclass.Count
        
        Set .Buckets = GenericArray.Build(.Size)
        Set .Head = New GenericBucket
        Set .Head.Before = .Head
        Set .Head.After = .Head
        
        Set oldBucket = myclass.Head.After
        
        Do While oldBucket Is myclass.Head = False
            
            Call LoadHash(Hash, oldBucket.HashCode, .Size)
            
            Do
                Set newBucket = .Buckets.ElementAt(Hash.Index)
                
                If newBucket Is Nothing Then
                    Set newBucket = New GenericBucket
                    Set .Buckets.ElementAt(Hash.Index) = newBucket
                    
                    With newBucket
                        Set .After = c.Head
                        Set .Before = c.Head.Before
                        Set .Before.After = newBucket
                        Set .After.Before = newBucket
                        .State = (HAS_ENTRY) '!
                        .HashCode = oldBucket.HashCode
                        Set .Key = oldBucket.Key: Set .Value = oldBucket.Value
                    End With
        
                    Exit Do
                End If
                
                If Not (newBucket.State And HAS_COLLISION) = HAS_COLLISION Then _
                    newBucket.State = (HAS_COLLISION Or HAS_ENTRY): Call InterlockedIncrement(myclass.Occupancy)
                
                Hash.Index = Modulo(Hash.Index + Hash.Incr, Hash.Size)
            Loop
            
            Set oldBucket = oldBucket.After
            
        Loop
    End With
    
    Set IGenericCollection_Copy = System.CreateInstance(New GenericLinkedMap, VarPtr(c), LenB(c))
   
End Function

Private Function IGenericCollection_Iterator() As IGenericIterator
    
    Dim c As Member
    
    Call LoadIterator(c.Iterator, myclass.Head)
    Set IGenericCollection_Iterator = System.CreateInstance(New GenericLinkedMap, VarPtr(c), LenB(c))
    
End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
  
    With myclass.Iterator
        If (.Bucket Is .Head = False) Then
            Set .current = .Bucket.ToPair
            Set Out = .current
            Set .Bucket = .Bucket.After
            IGenericIterator_HasNext = True
            Exit Function
        End If
       
        If .HasLoaded = False Then _
            Call System.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator
        
        Set .Bucket = Nothing
        Set .current = Nothing

    End With
    
End Function

Private Sub IGenericIterator_Reset()
    With myclass.Iterator
         If .HasLoaded = False Then _
            Call System.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
            
        Call LoadIterator(myclass.Iterator, .Head)
    End With
End Sub

