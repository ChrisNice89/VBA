VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericMap"
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
'* Class:  GenericMap
'*
'* Purpose: A GenericMap represents a data structure that can store objects as key value pairs.
'*          It is dynamic in nature means the size of the map is growing according to the need.
'*          You can search for a value in an instance of GenericMap class using the corresponding key.
'*          It holds both keys and values (A key cannot be null, but a value can be) in a
'*          pseudo random order determined by hashes calculated from keys.
'*          It is also sparse, with unoccupied holes in a pre-allocated array.
'*          A GenericMap object consists of buckets that contain the elements of the collection.
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
'*          A GenericMap doesn’t maintain the order of stored values.
'*          This map uses double hashing. There are hashsize buckets in the map, and ech bucket can contain 0 or 1 element.
'*
'*======================================================================================================================

'@Exposed
'@PredeclaredId
'@Folder "<T>Enumerable"
Option Explicit

Implements IGeneric
Implements IGenericDictionary
Implements IGenericCollection
Implements IGenericIterator

#If Win64 Then
    Private Const POINTERSIZE As LongPtr = 8
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dst As Any, ByRef Src As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
    Private Const POINTERSIZE As Long = 4
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dst As Any, ByRef Src As Any, ByVal Length As Long)
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal Length As Long)
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

Private Const MAXLONG As Long = 2 ^ 31 - 1
Private Const DEFAULT_LOADFACTOR As Double = 0.7
Private Const DEFAULT_CAPACITY As Long = 23 ' 23*LoadFactor(0.7) =16
Private Const S_OK As Long = &H0&

'Bucketstate
Private Const IS_EMPTY As Byte = 0
Private Const HAS_ENTRY As Byte = 2 ^ 0
Private Const HAS_COLLISION As Byte = 2 ^ 1

Private Type Constant
    MaxVal As Long
    InitialSize As Long
    Primes() As Variant
    HashPrime As Long
    RehashBarrier As Long
End Type

'
Private Type HashItem
    Incr As Long
    Size As Long
    Index As Long
    Code As Long
End Type

'
Private Type Bucket
    HashCode As Long
    Key As IGeneric
    Value As IGeneric
    STATE As Byte
End Type
'
Private Type Iterator
    Type As DictionaryData
    Index As Long
    Size As Long
    HasLoaded As Boolean
    Current As IGeneric
    Map As GenericMap
End Type
'
Private Type Member
    Const As Constant
    Occupancy As Long
    Count As Long
    RawSize As Long
    Size As Long
    LoadSize As Long
    LoadFactor As Double
    Buckets() As Bucket
    Iterator As Iterator
End Type
Private myclass As Member

Public Property Get LoadSize() As Long: LoadSize = myclass.LoadSize: End Property
Public Property Get LoadFactor() As Double: LoadFactor = myclass.LoadFactor: End Property
Public Property Get Count() As Long: Count = myclass.Count: End Property
Public Property Set Item(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value, True): End Property
Public Property Get Item(ByVal Key As IGeneric) As IGeneric
Attribute Item.VB_UserMemId = 0

    Dim i As Long: i = FindEntry(Key)
    
    If (i < 0) Then _
        Call System.Throw(Me, "Public Property Get Item").NoSuchElement
    
    Set Item = myclass.Buckets(i).Value: Exit Property

End Property
Public Function Contains(ByVal Key As IGeneric) As Boolean: Contains = FindEntry(Key) >= 0: End Function 'TryGetValue(Key, Nothing): End Function
Public Function GetKeys() As GenericArray

    ReDim Keys(GenericArray.LowerBound To myclass.Count - GenericArray.LowerBound + 1) As IGeneric
    Dim Index As Long

    Dim i As Long
    
    For i = LBound(myclass.Buckets) To UBound(myclass.Buckets)
        If (myclass.Buckets(i).STATE And HAS_ENTRY) = HAS_ENTRY Then _
            Set Keys(InterlockedIncrement(Index)) = myclass.Buckets(i).Key
    Next
    
    Set GetKeys = GenericArray.BuildFrom(Keys)

End Function

Public Function GetValues() As GenericArray

    ReDim Values(GenericArray.LowerBound To myclass.Count - GenericArray.LowerBound + 1) As IGeneric
    Dim Index As Long

    Dim i As Long
    
    For i = LBound(myclass.Buckets) To UBound(myclass.Buckets)
        If (myclass.Buckets(i).STATE And HAS_ENTRY) = HAS_ENTRY Then _
            Set Values(InterlockedIncrement(Index)) = myclass.Buckets(i).Value
    Next
    
    Set GetValues = GenericArray.BuildFrom(Values)
    
End Function

Private Function getMinPrime() As Long: getMinPrime = myclass.Const.Primes(LBound(myclass.Const.Primes, 1)): End Function
Private Function EmptyBucket() As Bucket: End Function
Private Function Modulo(ByVal a As Double, ByVal m As Double) As Long: Modulo = (a - (Int(a / m) * m)): End Function
Private Function Decrement(ByRef i As Long) As Long: i = (Not -i): Decrement = i: End Function
Private Function Increment(ByRef i As Long) As Long: i = (-(Not i)): Increment = i: End Function

Public Function Build( _
                        Optional ByVal Capacity As Long = DEFAULT_CAPACITY, _
                        Optional ByVal LoadFactor As Double = DEFAULT_LOADFACTOR) As GenericMap
    
    If (Me Is GenericMap) = False Then _
        Call System.Throw(Me, "Public Function Build").BuildExeption
    
    With myclass
        
        If (Capacity < 0) Then _
           Call System.Throw(Me, "Build").ArgOutOfRange("Capacity < 0")
        
        If Not (LoadFactor >= 0.1 And LoadFactor <= 1#) Then _
           Call System.Throw(Me, "Build").ArgOutOfRange("Not (LoadFactor >= 0.1 And LoadFactor <= 1#")
    
        With .Const
            .Primes = Array(3, 7, 11, 17, 23, 29, 37, 47, 59, 71, 89, 107, 131, 163, 197, 239, 293, 353, 431, 521, 631, 761, 919, _
                            1103, 1327, 1597, 1931, 2333, 2801, 3371, 4049, 4861, 5839, 7013, 8419, 10103, 12143, 14591, _
                            17519, 21023, 25229, 30293, 36353, 43627, 52361, 62851, 75431, 90523, 108631, 130363, 156437, _
                            187751, 225307, 270371, 324449, 389357, 467237, 560689, 672827, 807403, 968897, 1162687, 1395263, _
                            1674319, 2009191, 2411033, 2893249, 3471899, 4166287, 4999559, 5999471, 7199369)
        
            .MaxVal = MAXLONG
            .InitialSize = getMinPrime
            .HashPrime = 101
            .RehashBarrier = 100
        End With
        
        .LoadFactor = LoadFactor
        .RawSize = Capacity / .LoadFactor
        
        If (.RawSize > MAXLONG) Then _
           Call System.Throw(Me, "Build").Overflow(".RawSize > Skynet.MaxArraySize")
        
        .Size = IIf(.RawSize > .Const.InitialSize, getPrime(.RawSize), .Const.InitialSize)
        ReDim .Buckets(.Size - 1)
        
        .LoadSize = Int(.LoadFactor * .Size)
        
        If Not (.LoadSize < .Size) Then _
            Call System.Throw(Me, "Build").InvalidInput("Based on the current algorithm, loadSize must be less than Size.")
        
    End With
    
    Set Build = System.CreateInstance(New GenericMap, VarPtr(myclass), LenB(myclass))
    
End Function

Public Sub Clear()
    Erase myclass.Buckets
    ReDim myclass.Buckets(myclass.Size - 1)
    myclass.Count = 0
    myclass.Occupancy = 0
End Sub

Public Function TryByIndex(ByVal Index As Long, ByRef Key As IGeneric, ByRef Value As IGeneric) As Boolean
    With myclass.Buckets(Index)
        If (.STATE And HAS_ENTRY) = HAS_ENTRY Then
            Set Key = .Key
            Set Value = .Value
            TryByIndex = True
        End If
    End With
End Function

Public Function TryGetValue(ByVal Key As IGeneric, ByRef Value As IGeneric) As Boolean
             
    Dim i As Long: i = FindEntry(Key)
    
    If (i >= 0) Then _
        Set Value = myclass.Buckets(i).Value: TryGetValue = True
        
End Function

Private Function FindEntry(ByVal Key As IGeneric) As Long

    Dim Try As Long
    Dim Hash As HashItem
    
    If (Key Is Nothing) Then _
        Call System.Throw(Me, "Private Function FindEntry").IsNothing("(Key Is Nothing)")

    Call LoadHash(Hash, Key.HashValue, myclass.Size)
  
    Do
        With myclass.Buckets(Hash.Index)
            If ((.STATE And HAS_ENTRY) = HAS_ENTRY) Then _
                If .HashCode = Hash.Code Then _
                    If Key.Equals(.Key) Then _
                        FindEntry = Hash.Index: Exit Function

            If (Not (.STATE And HAS_COLLISION) = HAS_COLLISION) Then _
                FindEntry = -1: Exit Function
        
        End With
       
        If ((InterlockedIncrement(Try)) > Hash.Size) Then _
            FindEntry = -1: Exit Function

        Hash.Index = Modulo(Hash.Index + Hash.Incr, Hash.Size)
        
    Loop
    
    FindEntry = -1
   
End Function

Public Sub Remove(ByVal Key As IGeneric)
    
    Dim S As Byte
    Dim i As Long: i = FindEntry(Key)
    
    If (i >= 0) Then
        With myclass
            S = (.Buckets(i).STATE And Not HAS_ENTRY)
            .Buckets(i) = EmptyBucket: .Buckets(i).STATE = S
            Call InterlockedDecrement(.Count)
        End With
    End If
    
End Sub

Private Function getPrime(ByVal Min As Long) As Long
    
    Dim P As Long
    
    If Min < 0 Then _
        Call System.Throw(Me, "getPrime").InvalidInput("Min < 0")
    
    With myclass.Const
        For P = LBound(.Primes) To UBound(.Primes)
            If .Primes(P) >= Min Then _
                getPrime = .Primes(P): Exit Function
        Next
        
        For P = Min To .MaxVal
            If IsPrime(P) And ((P - 1) Mod .HashPrime) <> 0 Then _
                getPrime = P: Exit Function
        Next
    End With
    
    getPrime = Min
    
End Function

Private Function IsPrime(ByVal Candidate As Long) As Boolean
    
    Dim Limit As Long: Limit = Sqr(Candidate)
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

Private Function Expandprime(ByVal oldSize As Long) As Long
    
    Dim newSize As Long: newSize = oldSize * 2
    
    If newSize > myclass.Const.MaxVal Then
        Expandprime = myclass.Const.MaxVal
    Else
        Expandprime = getPrime(newSize)
    End If
    
End Function

Private Sub Expand()
    Call Rehash(Expandprime(myclass.Size), False)
End Sub

Private Sub LoadHash(ByRef h As HashItem, ByVal HashCode As Long, ByVal Size As Long)
    
    With h
        .Code = HashCode
        .Size = Size
        .Index = Modulo(.Code, .Size)
        .Incr = 1 + Modulo((CDbl(.Code) * myclass.Const.HashPrime), (.Size - 1))
    End With
    
End Sub

Private Sub Rehash(ByVal Size As Long, ByVal ForceNewHashing As Boolean)

    Dim NewBuckets() As Bucket: ReDim NewBuckets(Size - 1)
    Dim i As Long
    Dim Hash As HashItem
    
    With myclass
        .Occupancy = 0
        
        For i = LBound(.Buckets) To UBound(.Buckets)
            If (.Buckets(i).STATE And HAS_ENTRY) = HAS_ENTRY Then
                
                If Not ForceNewHashing Then
                    Call LoadHash(Hash, .Buckets(i).HashCode, Size)
                Else
                    Call LoadHash(Hash, .Buckets(i).Key.HashValue, Size)
                End If
              
                Do
                    With NewBuckets(Hash.Index)
                        If (.STATE = IS_EMPTY) Then
                            .STATE = HAS_ENTRY
                            .HashCode = Hash.Code
                            Set .Key = myclass.Buckets(i).Key
                            Set .Value = myclass.Buckets(i).Value
                            Exit Do
                        End If
                       
                        If Not (.STATE And HAS_COLLISION) = HAS_COLLISION Then _
                            .STATE = (HAS_ENTRY + HAS_COLLISION): Call InterlockedIncrement(myclass.Occupancy)
                    
                    End With
                    
                    Hash.Index = Modulo(Hash.Index + Hash.Incr, Hash.Size)
                Loop
            End If
        Next
    
        'Todo avoid Hardcopy NewBuckets!
        .Buckets = NewBuckets
        .LoadSize = Int(.LoadFactor * Size)
        .Size = Size
    End With
    
End Sub

Public Sub Add( _
                  ByVal Key As IGeneric, _
                  ByVal Value As IGeneric, _
                  Optional ByVal Overwrite As Boolean)

    Dim Try As Long
    Dim Hash As HashItem
    
    If (Key Is Nothing) Then _
        Call System.Throw(Me, "Public Function Add").IsNothing("(Key Is Nothing)")
        
    With myclass
    
        If (.Count >= .LoadSize) Then
            Call Rehash(Expandprime(.Size), False)
        ElseIf ((.Occupancy > .LoadSize) And (.Count > .Const.RehashBarrier)) Then
            Call Rehash(.Size, False)
        End If
        
        Call LoadHash(Hash, Key.HashValue, .Size)
   
        Do
            With myclass.Buckets(Hash.Index)
                If (Not (.STATE And HAS_ENTRY) = HAS_ENTRY) Then
                    .STATE = (HAS_ENTRY Or .STATE)
                    .HashCode = Hash.Code
                    Set .Key = Key: Set .Value = Value
                    Call InterlockedIncrement(myclass.Count)
                    Exit Sub
                End If
                
                If .HashCode = Hash.Code Then
                    If Key.Equals(.Key) Then
                        If (Overwrite) Then
                            Set .Value = Value
                        Else
                            Call System.Throw(Me, "Public Function Add").DuplicateNotAllowed(Key)
                        End If
                        Exit Sub
                    End If
                End If
                
                If (Not (.STATE And HAS_COLLISION) = HAS_COLLISION) Then _
                    .STATE = (HAS_ENTRY Or HAS_COLLISION): Call InterlockedIncrement(myclass.Occupancy)
                
            End With
            
            Hash.Index = Modulo(Hash.Index + Hash.Incr, Hash.Size)
        Loop While (InterlockedIncrement(Try) < Hash.Size)
    
    End With
    
End Sub

Private Sub Class_Terminate()
    Call System.Dispose(Me)
End Sub

Private Function CopyOf(ByVal t As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long, ByVal Caller As String) As GenericArray
    
    If (GArray Is Nothing) Then _
        Call System.Throw(Me, Caller).IsNothing("(GArray Is Nothing)")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call System.Throw(Me, Caller).IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < myclass.Count) Then _
        Call System.Throw(Me, Caller).InvalidInput("(GArray.Length - Index < myClass.Count)")
    
    Dim i As Long
    
    Select Case True
        Case t = KeyData
            For i = LBound(myclass.Buckets) To UBound(myclass.Buckets)
                If (myclass.Buckets(i).STATE And HAS_ENTRY) = HAS_ENTRY Then _
                    Set GArray(ArrayIndex) = myclass.Buckets(i).Key: Call InterlockedIncrement(ArrayIndex)
            Next
        Case t = ValueData
            For i = LBound(myclass.Buckets) To UBound(myclass.Buckets)
                If (myclass.Buckets(i).STATE And HAS_ENTRY) = HAS_ENTRY Then _
                    Set GArray(ArrayIndex) = myclass.Buckets(i).Value: Call InterlockedIncrement(ArrayIndex)
            Next
        Case Else
             For i = LBound(myclass.Buckets) To UBound(myclass.Buckets)
                If (myclass.Buckets(i).STATE And HAS_ENTRY) = HAS_ENTRY Then _
                    Set GArray(ArrayIndex) = GenericPair(myclass.Buckets(i).Key, myclass.Buckets(i).Value): Call InterlockedIncrement(ArrayIndex)
            Next
    End Select
    
    Set CopyOf = GArray
    
End Function

Public Function IteratorOf(ByVal t As DictionaryData) As IGenericIterator
                     
    Dim c As Member
    
    Call LoadIterator(c.Iterator, Me, t, myclass.Size)
    Set IteratorOf = System.CreateInstance(New GenericMap, VarPtr(c), LenB(c))
    
End Function

Private Sub LoadIterator( _
                            ByRef Iterator As Iterator, _
                            ByVal Map As GenericMap, _
                            ByVal t As DictionaryData, _
                            ByVal Size As Long)
    
    With Iterator
        .Type = t
        Set .Map = Map
        Set .Current = Nothing
        .Size = Size
        .Index = Size
        .HasLoaded = True
    End With
    
End Sub

'IGeneric
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GenericMap: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericMap):  End Property

Private Function IGeneric_Clone() As IGeneric
    
    Dim b As Bucket
    Dim c As Member
    c = myclass 'hardcopy
    
    Set IGeneric_Clone = System.CreateInstance(New GenericMap, VarPtr(c), LenB(c))
   
End Function

Private Sub IGeneric_Dispose()
    myclass = EmptyStructure
End Sub
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
                            "LoadFactor: " & .LoadFactor & vbNewLine & _
                            "LoadSize: " & .LoadSize & vbNewLine & _
                            "Elements: " & .Count & vbNewLine & _
                            "Occupancy: " & .Occupancy
    End With
End Function

'IGenericDictionary
Private Property Set IGenericDictionary_Item(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value, True): End Property
Private Property Get IGenericDictionary_Item(ByVal Key As IGeneric) As IGeneric: Set IGenericDictionary_Item = myclass.Buckets(FindEntry(Key)).Value: End Property
Private Property Get IGenericDictionary_Count() As Long: IGenericDictionary_Count = myclass.Count: End Property

Private Function IGenericDictionary_ContainsKey(ByVal Key As IGeneric) As Boolean: IGenericDictionary_ContainsKey = Me.TryGetValue(Key, Nothing): End Function
Private Function IGenericDictionary_ContainsValue(ByVal Value As IGeneric) As Boolean
        
        Dim i As Long
        
        If Not Value Is Nothing Then
            For i = LBound(myclass.Buckets) To UBound(myclass.Buckets)
                If (myclass.Buckets(i).STATE And HAS_ENTRY) = HAS_ENTRY Then _
                    If (Value.Equals(myclass.Buckets(i).Value)) Then _
                        IGenericDictionary_ContainsValue = True: Exit Function
            Next
            
        Else
            For i = LBound(myclass.Buckets) To UBound(myclass.Buckets)
                If (myclass.Buckets(i).STATE And HAS_ENTRY) = HAS_ENTRY Then _
                    If (myclass.Buckets(i).Value Is Nothing) Then _
                        IGenericDictionary_ContainsValue = True: Exit Function
            Next
        End If
            
End Function
Private Function IGenericDictionary_TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: IGenericDictionary_TryGetValue = Me.TryGetValue(Key, Value): End Function
Private Sub IGenericDictionary_Remove(ByVal Key As IGeneric): Call Me.Remove(Key): End Sub
Private Sub IGenericDictionary_Add(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value, False): End Sub
Private Sub IGenericDictionary_CopyOf(ByVal t As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(t:=t, GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="Private Sub IGenericDictionary_CopyOf"): End Sub
Private Sub IGenericDictionary_Clear(): Call Me.Clear: End Sub
Private Function IGenericDictionary_IteratorOf(ByVal t As DictionaryData) As IGenericIterator: Set IGenericDictionary_IteratorOf = Me.IteratorOf(t:=t): End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.Contains(Key:=Item): End Function
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(t:=PairData, GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="Private Sub IGenericCollection_CopyTo"): End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_ToArray() As GenericArray
    
    Dim GArray As GenericArray
    Set GArray = GenericArray.Build(myclass.Count)
    
    Call IGenericCollection_CopyTo(GArray:=GArray, ArrayIndex:=GArray.LowerBound)
    Set IGenericCollection_ToArray = GArray
    
End Function
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.IteratorOf(t:=PairData): End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.Current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    Dim b As Bucket
    
    With myclass.Iterator
        Do While InterlockedDecrement(.Index) > -1
            If .Map.TryByIndex(.Index, b.Key, b.Value) Then
               
                If .Type = KeyData Then
                    Set .Current = b.Key
        
                ElseIf .Type = ValueData Then
                    Set .Current = b.Value
                
                Else
                    Set .Current = GenericPair(b.Key, b.Value)
                
                End If
                
                Set Out = .Current
                IGenericIterator_HasNext = True
                Exit Function
            End If
        Loop
        
        If .HasLoaded = False Then _
            Call System.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator

        Set .Current = Nothing
        Set Out = Nothing
        
    End With
End Function

Private Sub IGenericIterator_Reset()
    With myclass.Iterator
         If .HasLoaded = False Then _
            Call System.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
            
        Call LoadIterator(myclass.Iterator, .Map, .Type, .Size)
    End With
End Sub

