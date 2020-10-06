VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@PredeclaredId
'@Folder "<T>Enumerable"
Option Explicit

Implements IPrototype
Implements IGeneric
Implements IGenericDictionary
Implements IGenericCollection
Implements IGenericIterator

Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long

Private Const DEFAULT_LOADFACTOR As Double = 0.7
Private Const DEFAULT_CAPACITY As Long = 23 ' 23*LoadFactor(0.7) =16

'Bucketstate
Private Const ZERO As Byte = 0
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
End Type

'
Private Type Bucket
    HashCode As Long
    Key As IGeneric
    Value As IGeneric
    State As Byte
End Type
'
Private Type Iterator
    Type As DictionaryData
    Index As Long
    HasLoaded As Boolean
    Current As IGeneric
    Buckets() As Bucket
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

'
Private myClass As Member
Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get LoadSize() As Long: LoadSize = myClass.LoadSize: End Property
Public Property Get LoadFactor() As Double: LoadFactor = myClass.LoadFactor: End Property
Public Property Get Count() As Long: Count = myClass.Count: End Property
Public Property Let Item(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value, True): End Property
Public Property Set Item(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value, True): End Property
Public Property Get Item(ByVal Key As IGeneric) As IGeneric: Set Item = myClass.Buckets(FindEntry(Key)).Value: End Property
Attribute Item.VB_UserMemId = 0
Public Function Contains(ByVal Key As IGeneric) As Boolean: Contains = FindEntry(Key) >= 0: End Function 'TryGetValue(Key, Nothing): End Function
Public Function GetKeys() As GenericArray

    ReDim Keys(GenericArray.LowerBound To myClass.Count) As IGeneric
    Dim Index As Long: Index = GenericArray.LowerBound

    Dim i As Long
    
    For i = LBound(myClass.Buckets) To UBound(myClass.Buckets)
        If (myClass.Buckets(i).State And HAS_ENTRY) = HAS_ENTRY Then _
            Set Keys(Index) = myClass.Buckets(i).Key: Call InterlockedIncrement(Index)
    Next
    
    Set GetKeys = GenericArray.BuildFrom(Keys)

End Function

Public Function GetValues() As GenericArray

    ReDim Values(GenericArray.LowerBound To myClass.Count) As IGeneric
    Dim Index As Long: Index = GenericArray.LowerBound

    Dim i As Long
    
    For i = LBound(myClass.Buckets) To UBound(myClass.Buckets)
        If (myClass.Buckets(i).State And HAS_ENTRY) = HAS_ENTRY Then _
            Set Values(Index) = myClass.Buckets(i).Value: Call InterlockedIncrement(Index)
    Next
    
    Set GetValues = GenericArray.BuildFrom(Values)
    
End Function

Private Function getMinPrime() As Long: getMinPrime = myClass.Const.Primes(LBound(myClass.Const.Primes, 1)): End Function
Private Function EmptyBucket() As Bucket: End Function
Private Function Modulo(ByVal A As Double, ByVal m As Double) As Long: Modulo = (A - (Int(A / m) * m)): End Function
Private Function Decrement(ByRef i As Long) As Long: i = (Not -i): Decrement = i: End Function
Private Function Increment(ByRef i As Long) As Long: i = (-(Not i)): Increment = i: End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Virtual.Size = LenB(myClass)
End Sub

Public Function Build( _
                        Optional ByVal Capacity As Long = DEFAULT_CAPACITY, _
                        Optional ByVal LoadFactor As Double = DEFAULT_LOADFACTOR) As GenericMap
    
    If (Me Is GenericMap) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption
    
    With myClass
        With .Const
            .Primes = Array(3, 7, 11, 17, 23, 29, 37, 47, 59, 71, 89, 107, 131, 163, 197, 239, 293, 353, 431, 521, 631, 761, 919, _
                            1103, 1327, 1597, 1931, 2333, 2801, 3371, 4049, 4861, 5839, 7013, 8419, 10103, 12143, 14591, _
                            17519, 21023, 25229, 30293, 36353, 43627, 52361, 62851, 75431, 90523, 108631, 130363, 156437, _
                            187751, 225307, 270371, 324449, 389357, 467237, 560689, 672827, 807403, 968897, 1162687, 1395263, _
                            1674319, 2009191, 2411033, 2893249, 3471899, 4166287, 4999559, 5999471, 7199369)
        
            .MaxVal = Skynet.MaxLong
            .InitialSize = getMinPrime
            .HashPrime = 101
            .RehashBarrier = 100
        End With
        
        If (Capacity < 0) Then _
           Call Skynet.Throw(Me, "Build").ArgOutOfRange("Capacity < 0")
        
        If Not (LoadFactor >= 0.1 And LoadFactor <= 1#) Then _
           Call Skynet.Throw(Me, "Build").ArgOutOfRange("Not (LoadFactor >= 0.1 And LoadFactor <= 1#")
    
        .LoadFactor = LoadFactor
        .RawSize = Capacity / .LoadFactor
        
        If (.RawSize > Skynet.MaxLong) Then _
           Call Skynet.Throw(Me, "Build").Overflow(".RawSize > Skynet.MaxArraySize")
        
        .Size = IIf(.RawSize > .Const.InitialSize, getPrime(.RawSize), .Const.InitialSize)
        ReDim .Buckets(.Size - 1)
        
        .LoadSize = Int(.LoadFactor * .Size)
        If Not (.LoadSize < .Size) Then _
            Call Skynet.Throw(Me, "Build").InvalidInput("Based on the current algorithm, loadSize must be less than Size.")
            
    End With
    
    Set Build = Skynet.CreateInstance(New GenericMap, Virtual.Pointer, Virtual.Size)
    
End Function

Public Sub Clear()
    Erase myClass.Buckets
    ReDim myClass.Buckets(myClass.Size - 1)
    myClass.Count = 0
    myClass.Occupancy = 0
End Sub

Public Function TryGetValue(ByVal Key As IGeneric, ByRef Value As IGeneric) As Boolean
             
    Dim Try As Long
    Dim Hash As HashItem
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function TryGetValue").KeyIsNothing
    
    Call LoadHash(Hash, Key.HashValue, myClass.Size)
    With Hash
        Do
            With myClass.Buckets(.Index)
                If Not (.Key Is Nothing) Then _
                    If .Key.Equals(Key) Then _
                        Set Value = .Value: TryGetValue = True: Exit Function
             
                If (Not (.State And HAS_COLLISION) = HAS_COLLISION) Then _
                    Exit Function
            
            End With
            If ((InterlockedIncrement(Try)) > .Size) Then _
                Exit Function
                
            .Index = Modulo(.Index + .Incr, .Size)
        Loop
    End With
    
End Function

Private Function FindEntry(ByVal Key As IGeneric) As Long

    Dim Try As Long
    Dim Hash As HashItem
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "FindEntry").KeyIsNothing
    
    Call LoadHash(Hash, Key.HashValue, myClass.Size)
  
    Do
        With myClass.Buckets(Hash.Index)
            If Not .Key Is Nothing Then _
                If .Key.Equals(Key) Then _
                    FindEntry = Hash.Index: Exit Function

            If (Not (.State And HAS_COLLISION) = HAS_COLLISION) Then _
                FindEntry = -1: Exit Function
        
        End With
        With Hash
            If ((InterlockedIncrement(Try)) > .Size) Then _
                FindEntry = -1: Exit Function
    
            .Index = Modulo(.Index + .Incr, .Size)
        End With
    Loop
    
    FindEntry = -1
   
End Function

Public Sub Remove(ByVal Key As IGeneric)
    
    Dim s As Byte
    Dim i As Long: i = FindEntry(Key)
    
    If (i >= 0) Then
        With myClass
            s = (.Buckets(i).State And Not HAS_ENTRY)
            .Buckets(i) = EmptyBucket: .Buckets(i).State = s
            Call InterlockedDecrement(.Count)
        End With
    End If
End Sub

Private Function getPrime(ByVal Min As Long) As Long
    
    Dim p As Long
    
    If Min < 0 Then _
        Call Skynet.Throw(Me, "getPrime").InvalidInput("Min < 0")
    
    With myClass.Const
        For p = LBound(.Primes) To UBound(.Primes)
            If .Primes(p) >= Min Then _
                getPrime = .Primes(p): Exit Function
        Next
        
        For p = Min To .MaxVal
            If IsPrime(p) And ((p - 1) Mod .HashPrime) <> 0 Then _
                getPrime = p: Exit Function
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
    
    If newSize > myClass.Const.MaxVal Then
        Expandprime = myClass.Const.MaxVal
    Else
        Expandprime = getPrime(newSize)
    End If
    
End Function

Private Sub Expand()
    Call Rehash(Expandprime(myClass.Size), False)
End Sub

Private Sub LoadHash(ByRef H As HashItem, ByVal HashValue As Long, ByVal Size As Long)
    
    With H
        .Size = Size
        .Index = Modulo(HashValue, .Size)
        .Incr = 1 + Modulo((CDbl(HashValue) * myClass.Const.HashPrime), (.Size - 1))
    End With
    
End Sub

Private Sub Rehash(ByVal Size As Long, ByVal ForceNewHashing As Boolean)

    Dim NewBuckets() As Bucket: ReDim NewBuckets(Size - 1)
    Dim i As Long
  
    With myClass
        .Occupancy = 0
        
        If ForceNewHashing Then
            For i = LBound(.Buckets) To UBound(.Buckets)
                If Not (.Buckets(i).Key Is Nothing) Then _
                    Call PutEntry(NewBuckets, .Buckets(i), .Buckets(i).Key.HashValue, Size)
            Next
        Else
            For i = LBound(.Buckets) To UBound(.Buckets)
                If Not (.Buckets(i).Key Is Nothing) Then _
                    Call PutEntry(NewBuckets, .Buckets(i), .Buckets(i).HashCode, Size)
            Next
        End If
        
        .Buckets = NewBuckets
        .LoadSize = Int(.LoadFactor * Size)
        .Size = Size
    End With
    
End Sub

Private Sub PutEntry( _
                        ByRef Buckets() As Bucket, _
                        ByRef Entry As Bucket, _
                        ByVal HashCode As Long, _
                        ByVal Size As Long)

    Dim Hash As HashItem
    Call LoadHash(Hash, HashCode, Size)
    Entry.State = HAS_ENTRY
    
    With Hash
        Do
            If (Buckets(.Index).Key Is Nothing) Then _
                Buckets(.Index) = Entry: Exit Sub
            
            If Not (Buckets(.Index).State And HAS_COLLISION) = HAS_COLLISION Then _
                Buckets(.Index).State = (HAS_ENTRY + HAS_COLLISION): Call InterlockedIncrement(myClass.Occupancy)
                
            .Index = Modulo(.Index + .Incr, .Size)
        Loop
    End With

End Sub

Public Sub Add( _
                  ByVal Key As IGeneric, _
                  ByVal Value As IGeneric, _
                  Optional ByVal Overwrite As Boolean)

    Dim Try As Long
    Dim HashValue As Long
    Dim Hash As HashItem
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Add").KeyIsNothing
        
    With myClass
        If (.Count >= .LoadSize) Then
            Call Rehash(Expandprime(.Size), False)
        ElseIf ((.Occupancy > .LoadSize) And (.Count > .Const.RehashBarrier)) Then
            Call Rehash(.Size, False)
        End If
        
        HashValue = Key.HashValue
        Call LoadHash(Hash, HashValue, .Size)
    End With
    
    With Hash
        Do
            With myClass.Buckets(.Index)
                If .Key Is Nothing Then
                    .State = HAS_ENTRY
                    .HashCode = HashValue
                    Set .Key = Key: Set .Value = Value
                    Call InterlockedIncrement(myClass.Count)
                    Exit Sub
                End If
                
                If .Key.Equals(Key) Then
                    If (Overwrite = False) Then _
                        Call Skynet.Throw(Me, "Public Function Add").DuplicateNotAllowed(Key)
        
                    Set .Value = Value: Exit Sub
                End If
                
                If (Not (.State And HAS_COLLISION) = HAS_COLLISION) Then _
                    .State = (HAS_ENTRY Or HAS_COLLISION): Call InterlockedIncrement(myClass.Occupancy)
                
            End With
            
            .Index = Modulo(.Index + .Incr, .Size)
        Loop While (InterlockedIncrement(Try) < .Size)
    End With

End Sub

Private Sub Class_Terminate()
    Call Skynet.Dispose(Me)
End Sub

Private Function CopyOf(ByVal T As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long, ByVal Caller As String) As GenericArray
    
    If (GArray Is Nothing) Then _
        Call Skynet.Throw(Me, Caller).NullPointer("GArray Is Nothing")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call Skynet.Throw(Me, Caller).IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < myClass.Count) Then _
        Call Skynet.Throw(Me, Caller).InvalidInput("(GArray.Length - Index < myClass.Count)")
    
    Dim i As Long
    
    Select Case True
        Case T = keys_
            For i = LBound(myClass.Buckets) To UBound(myClass.Buckets)
                If (myClass.Buckets(i).State And HAS_ENTRY) = HAS_ENTRY Then _
                    Set GArray(ArrayIndex) = myClass.Buckets(i).Key: Call InterlockedIncrement(ArrayIndex)
            Next
        Case T = Values_
            For i = LBound(myClass.Buckets) To UBound(myClass.Buckets)
                If (myClass.Buckets(i).State And HAS_ENTRY) = HAS_ENTRY Then _
                    Set GArray(ArrayIndex) = myClass.Buckets(i).Value: Call InterlockedIncrement(ArrayIndex)
            Next
        Case Else
             For i = LBound(myClass.Buckets) To UBound(myClass.Buckets)
                If (myClass.Buckets(i).State And HAS_ENTRY) = HAS_ENTRY Then _
                    Set GArray(ArrayIndex) = GenericPair(myClass.Buckets(i).Key, myClass.Buckets(i).Value): Call InterlockedIncrement(ArrayIndex)
            Next
    End Select
    
    Set CopyOf = GArray
    
End Function

Public Function Iterator(ByVal T As DictionaryData) As IGenericIterator
    
    Dim c As Member
    Dim ptr As LongPtr
    
    Call Skynet.CopyData(VarPtr(c), Virtual.Pointer, Virtual.Size)
    Call LoadIterator(c.Iterator, Me, T, myClass.Size) '! Order is important because skynet overwrites class structure

    Set Iterator = Skynet.CreateInstance(New GenericMap, VarPtr(c), Virtual.Size)
    
End Function

Private Sub LoadIterator( _
                            ByRef Iterator As Iterator, _
                            ByVal Map As GenericMap, _
                            ByVal T As DictionaryData, _
                            ByVal Size As Long)
    
    With Iterator
        .Type = T
        Set .Map = Map
        Set .Current = Nothing
        .Index = Size
        .HasLoaded = True
    End With
    
End Sub
'IGeneric
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = IGeneric_HashValue = ObjPtr(Me): End Property

Private Function IGeneric_Clone() As IGeneric
    
    Dim B As Bucket
    Dim c As Member
    c = myClass 'hardcopy
    
    Set IGeneric_Clone = Skynet.CreateInstance(New GenericMap, VarPtr(c), LenB(c))
   
End Function

Private Sub IGeneric_Dispose()
    If myClass.Iterator.HasLoaded = True Then
        Set myClass.Iterator.Map = Nothing
        Call Skynet.ClearData(Virtual.Pointer, Virtual.Size)
    Else
        myClass = EmptyStructure
    End If
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    IGeneric_Equals = (Me Is Other)
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    Call Skynet.Throw(Me, "IObject_CompareTo").NotSupported
End Function

Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericMap): End Function
Private Function IGeneric_ToString() As String
    With myClass
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
Private Property Get IGenericDictionary_Item(ByVal Key As IGeneric) As IGeneric: Set IGenericDictionary_Item = myClass.Buckets(FindEntry(Key)).Value: End Property
Private Property Get IGenericDictionary_Count() As Long: IGenericDictionary_Count = myClass.Count: End Property

Private Function IGenericDictionary_ContainsKey(ByVal Key As IGeneric) As Boolean: IGenericDictionary_ContainsKey = Me.TryGetValue(Key, Nothing): End Function
Private Function IGenericDictionary_ContainsValue(ByVal Value As IGeneric) As Boolean
        
        Dim i As Long
        
        If Not Value Is Nothing Then
            For i = LBound(myClass.Buckets) To UBound(myClass.Buckets)
                If (myClass.Buckets(i).State And HAS_ENTRY) = HAS_ENTRY Then _
                    If (Value.Equals(myClass.Buckets(i).Value)) Then _
                        IGenericDictionary_ContainsValue = True: Exit Function
            Next
            
        Else
            For i = LBound(myClass.Buckets) To UBound(myClass.Buckets)
                If (myClass.Buckets(i).State And HAS_ENTRY) = HAS_ENTRY Then _
                    If (myClass.Buckets(i).Value Is Nothing) Then _
                        IGenericDictionary_ContainsValue = True: Exit Function
            Next
        End If
            
End Function
Private Function IGenericDictionary_TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: IGenericDictionary_TryGetValue = Me.TryGetValue(Key, Value): End Function
Private Sub IGenericDictionary_Remove(ByVal Key As IGeneric): Call Me.Remove(Key): End Sub
Private Sub IGenericDictionary_Add(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value, False): End Sub
Private Sub IGenericDictionary_CopyOf(ByVal T As DictionaryData, ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(T:=T, GArray:=GArray, ArrayIndex:=ArrayIndex, Caller:="Private Sub IGenericDictionary_CopyOf"): End Sub
Private Sub IGenericDictionary_Clear(): Call Me.Clear: End Sub
Private Function IGenericDictionary_Iterator(ByVal T As DictionaryData) As IGenericIterator: Set IGenericDictionary_Iterator = Me.Iterator(T:=T): End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = False: End Property
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.Contains(Key:=Item): End Function
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call CopyOf(Pairs_, GArray, ArrayIndex, "Private Sub IGenericCollection_CopyTo"): End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_ToArray() As GenericArray
    
    Dim GArray As GenericArray
    Set GArray = GenericArray.Build(myClass.Count)
    
    Call IGenericCollection_CopyTo(GArray:=GArray, ArrayIndex:=GArray.LowerBound)
    Set IGenericCollection_ToArray = GArray
    
End Function
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Iterator(Pairs_): End Function

'
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myClass.Iterator.Current: End Property

Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    With myClass.Iterator
        If Not .Map.Count = 0 Then 'Keep Map in focus
            Do While .Index > 0
                If Not (myClass.Buckets(Decrement(.Index)).Key Is Nothing) Then
                    
                    If .Type = keys_ Then
                         Set .Current = myClass.Buckets(.Index).Key
                    
                    ElseIf .Type = Values_ Then
                        Set .Current = myClass.Buckets(.Index).Value
                    
                    Else
                        Set .Current = GenericPair(myClass.Buckets(.Index).Key, myClass.Buckets(.Index).Value)
                    End If
                    Set Out = .Current
                    IGenericIterator_HasNext = True
                    Exit Function
                End If
            Loop
        End If
        
        Set .Current = Nothing
        Set Out = Nothing
        
    End With
End Function

Private Sub IGenericIterator_Reset()
    Call LoadIterator(myClass.Iterator, myClass.Iterator.Map, myClass.Iterator.Type, UBound(myClass.Buckets))
End Sub
'IPrototype
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericMap: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GenericMap: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.Size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property

