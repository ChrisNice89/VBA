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
Implements IGenericCollection

'Implements IGenericEnumerator

Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
Private Const DEFAULTCAPACITY As Long = 32

'Bucketstate
Private Const ZERO As Byte = 0
Private Const IS_EMPTY As Byte = 0
Private Const HAS_ENTRY As Byte = 2 ^ 0
Private Const HAS_COLLISION As Byte = 2 ^ 1

Private Enum Data
    Values
    keys
End Enum

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
Private Type Member
    Const As Constant
    Occupancy As Long
    Count As Long
    RawSize As Long
    HashSize As Long
    LoadSize As Long
    LoadFactor As Double
    Buckets() As Bucket
    IsReadOnly As Boolean
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
Public Property Let Item(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value): End Property
Public Property Set Item(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value): End Property
Public Property Get Item(ByVal Key As IGeneric) As IGeneric: Set Item = myClass.Buckets(FindEntry(Key)).Value: End Property
Attribute Item.VB_UserMemId = 0
Public Function Contains(ByVal Key As IGeneric) As Boolean: Contains = FindEntry(Key) >= 0: End Function 'TryGetValue(Key, Nothing): End Function
Public Function GetKeys() As IGenericReadOnlyList: Set GetKeys = LoadKeys(GenericArray.Build(Me.Count), GenericArray.LowerBound).AsReadOnly: End Function
Public Function GetValues() As IGenericReadOnlyList: Set GetValues = LoadValues(GenericArray.Build(Me.Count), GenericArray.LowerBound).AsReadOnly: End Function

Private Function getMinPrime() As Long: getMinPrime = myClass.Const.Primes(LBound(myClass.Const.Primes, 1)): End Function
Private Function EmptyBucket() As Bucket: End Function
Private Function Modulo(ByVal a As Double, ByVal M As Double) As Long: Modulo = (a - (Int(a / M) * M)): End Function
Private Function Decrement(ByRef i As Long) As Long: i = (Not -i): Decrement = i: End Function
Private Function Increment(ByRef i As Long) As Long: i = (-(Not i)): Increment = i: End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
End Sub

Public Function Build( _
                        Optional ByVal Capacity As Long = DEFAULTCAPACITY, _
                        Optional ByVal LoadFactor As Double = 0.72) As GenericMap
    
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
           Call Skynet.Push(Me, "Build").ArgOutOfRange.Raise("Capacity < 0")
        
        If Not (LoadFactor >= 0.1 And LoadFactor <= 1#) Then _
           Call Skynet.Push(Me, "Build").ArgOutOfRange.Raise("Not (LoadFactor >= 0.1 And LoadFactor <= 1#")
    
        .LoadFactor = LoadFactor
        .RawSize = Capacity / .LoadFactor
        
        If (.RawSize > Skynet.MaxLong) Then _
           Call Skynet.Push(Me, "Build").Overflow.Raise(".RawSize > Skynet.MaxArraySize")
        
        .HashSize = IIf(.RawSize > .Const.InitialSize, getPrime(.RawSize), .Const.InitialSize)
        ReDim .Buckets(.HashSize - 1)
        
        .LoadSize = Int(.LoadFactor * .HashSize)
        If Not (.LoadSize < .HashSize) Then _
            Call Skynet.Push(Me, "Build").InvalidInput.Raise("Based on the current algorithm, loadsize must be less than hashsize.")
            
    End With
    
    Set Build = Skynet.Create(Me)
    
End Function

Private Function LoadKeys(ByVal keys As GenericArray, ByVal ArrayIndex As Long) As GenericArray
    Dim i As Long
    
    For i = LBound(myClass.Buckets) To UBound(myClass.Buckets)
        If (myClass.Buckets(i).State And HAS_ENTRY) = HAS_ENTRY Then _
            Set keys(ArrayIndex) = myClass.Buckets(i).Value: Call InterlockedIncrement(ArrayIndex)
    Next
    
    Set LoadKeys = keys
    
End Function

Private Function LoadValues(ByVal Values As GenericArray, ByVal ArrayIndex As Long) As GenericArray
    Dim i As Long
    
    For i = LBound(myClass.Buckets) To UBound(myClass.Buckets)
        If (myClass.Buckets(i).State And HAS_ENTRY) = HAS_ENTRY Then _
            Set Values(ArrayIndex) = myClass.Buckets(i).Value: Call InterlockedIncrement(ArrayIndex)
    Next
    
    Set LoadValues = Values
    
End Function

Public Function TryGetValue(ByVal Key As IGeneric, ByRef Value As IGeneric) As Boolean
             
    Dim Try As Long
    
    If (Key Is Nothing) Then _
        Call Skynet.Push(Me, "TryGetValue").NullPointer.Raise("Key is nothing")
    
    With InitialHashing(Key.HashValue, myClass.HashSize)
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

Public Function GetValueOrDefault(ByVal Key As IGeneric) As IGeneric
    
    Dim i As Long: i = FindEntry(Key)
    
    If (i) >= 0 Then _
        Set GetValueOrDefault = myClass.Buckets(i).Value: Exit Function
    
    Set GetValueOrDefault = Nothing 'Default
    
End Function

Private Function FindEntry(ByVal Key As IGeneric) As Long

    Dim Try As Long
    Dim h As HashItem
    
    If (Key Is Nothing) Then _
        Call Skynet.Push(Me, "FindEntry").NullPointer.Raise("Key is nothing")
    
    h = InitialHashing(Key.HashValue, myClass.HashSize)
    Do
        With myClass.Buckets(h.Index)
            If Not .Key Is Nothing Then _
                If .Key.Equals(Key) Then _
                    FindEntry = h.Index: Exit Function

            If (Not (.State And HAS_COLLISION) = HAS_COLLISION) Then _
                FindEntry = -1: Exit Function
        
        End With
        With h
            If ((InterlockedIncrement(Try)) > .Size) Then _
                FindEntry = -1: Exit Function
    
            .Index = Modulo(.Index + .Incr, .Size)
        End With
    Loop
    
    FindEntry = -1
   
End Function

Public Function Remove(ByVal Key As IGeneric) As Boolean
    
    Dim S As Byte
    Dim i As Long: i = FindEntry(Key)
    
    If (i >= 0) Then
        With myClass
            S = (.Buckets(i).State And Not HAS_ENTRY)
            .Buckets(i) = EmptyBucket: .Buckets(i).State = S
            Call InterlockedDecrement(.Count)
        End With
    End If
End Function

Public Sub Clear()
     Call Skynet.Generic(Me).Dispose
End Sub

Private Function getPrime(ByVal Min As Long) As Long
    
    Dim p As Long
    
    If Min < 0 Then _
        Call Skynet.Push(Me, "getPrime").InvalidInput.Raise("Min < 0")
    
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
    Call Rehash(Expandprime(myClass.HashSize), False)
End Sub

Private Function InitialHashing(ByVal HashValue As Long, ByVal HashSize As Long) As HashItem

    Dim hash As HashItem
    
    With hash
        .Size = HashSize
        .Index = Modulo(HashValue, .Size)
        .Incr = 1 + Modulo((CDbl(HashValue) * myClass.Const.HashPrime), (.Size - 1))
    End With
    
    LSet InitialHashing = hash

End Function

Private Sub Rehash(ByVal Size As Long, ByVal ForceNewHashing As Boolean)

    Dim NewBuckets() As Bucket: ReDim NewBuckets(Size - 1)
    Dim b As Bucket
    Dim i As Long
  
    With myClass
        .Occupancy = 0
        
        If ForceNewHashing Then
            For i = LBound(.Buckets) To UBound(.Buckets)
                If Not (.Buckets(i).Key Is Nothing) Then _
                    b = .Buckets(i): b.HashCode = b.Key.HashValue: Call PutEntry(NewBuckets, b, InitialHashing(b.HashCode, Size))
            Next
        Else
            For i = LBound(.Buckets) To UBound(.Buckets)
                If Not (.Buckets(i).Key Is Nothing) Then _
                    b = .Buckets(i): Call PutEntry(NewBuckets, b, InitialHashing(b.HashCode, Size))
            Next
        End If
        
        b = EmptyBucket
        .Buckets = NewBuckets
        .LoadSize = Int(.LoadFactor * Size)
        .HashSize = Size
    End With
    
End Sub

Private Sub PutEntry(ByRef Buckets() As Bucket, ByRef Entry As Bucket, ByRef hash As HashItem)

    Entry.State = HAS_ENTRY
    
    With hash
        Do
            If (Buckets(.Index).Key Is Nothing) Then _
                Buckets(.Index) = Entry: Exit Sub
            
            If Not (Buckets(.Index).State And HAS_COLLISION) = HAS_COLLISION Then _
                Buckets(.Index).State = (HAS_ENTRY + HAS_COLLISION): Call InterlockedIncrement(myClass.Occupancy)
                
            .Index = Modulo(.Index + .Incr, .Size)
        Loop
    End With

End Sub

Public Sub Add(ByRef Key As IGeneric, ByRef Value As IGeneric)

    Dim Try As Long
    Dim HashValue As Long
    
    If (Key Is Nothing) Then _
        Call Skynet.Push(Me, "Insert").NullPointer.Raise("Key is nothing")
        
    With myClass
        If (.Count >= .LoadSize) Then
            Call Rehash(Expandprime(.HashSize), False)
        ElseIf ((.Occupancy > .LoadSize) And (.Count > .Const.RehashBarrier)) Then
            Call Rehash(.HashSize, False)
        End If
        
        HashValue = Key.HashValue
            
    End With
    
    With InitialHashing(HashValue, myClass.HashSize)
        Do
            With myClass.Buckets(.Index)
                If .Key Is Nothing Then
                    .State = HAS_ENTRY
                    .HashCode = HashValue
                    Set .Key = Key: Set .Value = Value
                    Call InterlockedIncrement(myClass.Count)
                    Exit Sub
                End If
                
                If .Key.Equals(Key) Then _
                    Set .Value = Value: Exit Sub
             
                If (Not (.State And HAS_COLLISION) = HAS_COLLISION) Then _
                    .State = (HAS_ENTRY Or HAS_COLLISION): Call InterlockedIncrement(myClass.Occupancy)
                
            End With
            
            .Index = Modulo(.Index + .Incr, .Size)
        Loop While (InterlockedIncrement(Try) < .Size)
    End With

End Sub

Public Sub CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long)

    If (GArray Is Nothing) Then _
        Call Skynet.Push(Me, "CopyTo").NullPointer.Raise("GArray Is Nothing")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call Skynet.Push(Me, "CopyTo").IndexOutofRange.Raise("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < Me.Count) Then _
        Call Skynet.Push(Me, "CopyTo").InvalidInput.Raise("(GArray.Length - Index < Me.Count)")
    
    Call LoadValues(Values:=GArray, ArrayIndex:=ArrayIndex)

End Sub

Private Sub Class_Terminate()
    Call Skynet.Generic(Me).Dispose
End Sub

Private Property Get IGenericEnumerator_Current() As IGeneric

End Property

Private Function IGenericEnumerator_MoveNext() As Boolean
'    With myClass.Cache
'        Do While Increment(.CurrentSlot) < .Slots
'            If Not (.Buckets(.CurrentSlot).Key Is Nothing) Then
'                With .Buckets(.CurrentSlot)
''                    Set Key = .Key
''                    Set Value = .Value
'                End With
'                IEnumerator_MoveNext = True: Exit Function
'            End If
'        Loop
'
'    End With
End Function

Private Sub IGenericEnumerator_Reset()

End Sub
'IGeneric
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = IGeneric_HashValue = ObjPtr(Me): End Property
Private Function IGeneric_Clone() As IGeneric
    
    Dim b As Bucket
    Dim c As Member
    c = myClass
    'ReDim c.Buckets(myClass.HashSize)
    'Call Skynet.CopyData(ByVal VarPtr(c.Buckets(0)), ByVal VarPtr(myClass.Buckets(0)), LenB(b) * myClass.HashSize, False)
    Virtual.Pointer = VarPtr(c)
    Set IGeneric_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function

Private Sub IGeneric_Dispose()
   myClass = EmptyStructure
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    IGeneric_Equals = (Me Is Other)
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    Call Skynet.Push(Me, "IObject_CompareTo").NotSupported.Raise
End Function

Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericMap): End Function
Private Function IGeneric_ToString() As String
    With myClass
        IGeneric_ToString = _
                            "Class: " & TypeName$(Me) & vbNewLine & _
                            "HashSize: " & .HashSize & vbNewLine & _
                            "LoadFactor: " & .LoadFactor & vbNewLine & _
                            "LoadSize: " & .LoadSize & vbNewLine & _
                            "Elements: " & .Count & vbNewLine & _
                            "Occupancy: " & .Occupancy
    End With
End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = False: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = myClass.IsReadOnly: End Property
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call Me.CopyTo(GArray:=GArray, ArrayIndex:=ArrayIndex): End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = LoadValues(GenericArray.Build(Me.Count), GenericArray.LowerBound): End Function
Private Function IGenericCollection_NewEnumerator() As IGenericEnumerator

End Function

Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericMap: End Function
Private Property Get IPrototype_Default() As Object: Set IPrototype_Default = GenericMap: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property


