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

'Implements IGenericEnumerator

Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
Private Const DEFAULT_CAPACITY As Long = 16

'Bucketstate
Private Const ZERO As Byte = 0
Private Const IS_EMPTY As Byte = 0
Private Const HAS_ENTRY As Byte = 2 ^ 0
Private Const HAS_COLLISION As Byte = 2 ^ 1

Private Enum SealedClass
    SealedClass = 2 ^ 0
    Values = 2 ^ 1
    Keys = 2 ^ 2
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
    SealedClass As SealedClass
    isSealed As Boolean
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
Public Function GetKeys() As IGenericReadOnlyList: Set GetKeys = GenericReadOnlyList.Build(Internal:=AddKeysToArray(GenericArray.Build(myClass.Count), GenericArray.LowerBound), Size:=myClass.Count): End Function
Public Function GetValues() As IGenericReadOnlyList: Set GetValues = GenericReadOnlyList.Build(Internal:=AddValuesToArray(GenericArray.Build(myClass.Count), GenericArray.LowerBound), Size:=myClass.Count): End Function
Public Function ToArray() As IGenericReadOnlyList: Set ToArray = GenericReadOnlyList.Build(Internal:=AddPairsToArray(GenericArray.Build(myClass.Count), GenericArray.LowerBound), Size:=myClass.Count): End Function

Private Function getMinPrime() As Long: getMinPrime = myClass.Const.Primes(LBound(myClass.Const.Primes, 1)): End Function
Private Function EmptyBucket() As Bucket: End Function
Private Function Modulo(ByVal a As Double, ByVal m As Double) As Long: Modulo = (a - (Int(a / m) * m)): End Function
Private Function Decrement(ByRef i As Long) As Long: i = (Not -i): Decrement = i: End Function
Private Function Increment(ByRef i As Long) As Long: i = (-(Not i)): Increment = i: End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
End Sub

Public Function Build( _
                        Optional ByVal Capacity As Long = DEFAULT_CAPACITY, _
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
           Call Skynet.Throw(Me, "Build").ArgOutOfRange("Capacity < 0")
        
        If Not (LoadFactor >= 0.1 And LoadFactor <= 1#) Then _
           Call Skynet.Throw(Me, "Build").ArgOutOfRange("Not (LoadFactor >= 0.1 And LoadFactor <= 1#")
    
        .LoadFactor = LoadFactor
        .RawSize = Capacity / .LoadFactor
        
        If (.RawSize > Skynet.MaxLong) Then _
           Call Skynet.Throw(Me, "Build").Overflow(".RawSize > Skynet.MaxArraySize")
        
        .HashSize = IIf(.RawSize > .Const.InitialSize, getPrime(.RawSize), .Const.InitialSize)
        ReDim .Buckets(.HashSize - 1)
        
        .LoadSize = Int(.LoadFactor * .HashSize)
        If Not (.LoadSize < .HashSize) Then _
            Call Skynet.Throw(Me, "Build").InvalidInput("Based on the current algorithm, loadsize must be less than hashsize.")
            
    End With
    
    Set Build = Skynet.Create(Me)
    
End Function

Private Function AddPairsToArray(ByVal GArray As GenericArray, ByVal ArrayIndex As Long) As GenericArray
    Dim i As Long
    
    For i = LBound(myClass.Buckets) To UBound(myClass.Buckets)
        If (myClass.Buckets(i).State And HAS_ENTRY) = HAS_ENTRY Then _
            Set GArray(ArrayIndex) = GenericPair(myClass.Buckets(i).Key, myClass.Buckets(i).Value): Call InterlockedIncrement(ArrayIndex)
    Next
    
    Set AddPairsToArray = GArray
    
End Function

Private Function AddKeysToArray(ByVal GArray As GenericArray, ByVal ArrayIndex As Long) As GenericArray
    Dim i As Long
    
    For i = LBound(myClass.Buckets) To UBound(myClass.Buckets)
        If (myClass.Buckets(i).State And HAS_ENTRY) = HAS_ENTRY Then _
            Set GArray(ArrayIndex) = myClass.Buckets(i).Value: Call InterlockedIncrement(ArrayIndex)
    Next
    
    Set AddKeysToArray = GArray
    
End Function

Private Function AddValuesToArray(ByVal GArray As GenericArray, ByVal ArrayIndex As Long) As GenericArray
    Dim i As Long
    
    For i = LBound(myClass.Buckets) To UBound(myClass.Buckets)
        If (myClass.Buckets(i).State And HAS_ENTRY) = HAS_ENTRY Then _
            Set GArray(ArrayIndex) = myClass.Buckets(i).Value: Call InterlockedIncrement(ArrayIndex)
    Next
    
    Set AddValuesToArray = GArray
    
End Function

Public Function TryGetValue(ByVal Key As IGeneric, ByRef Value As IGeneric) As Boolean
             
    Dim Try As Long
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function TryGetValue").KeyIsNothing
    
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

Private Function FindEntry(ByVal Key As IGeneric) As Long

    Dim Try As Long
    Dim h As HashItem
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "FindEntry").KeyIsNothing
    
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

Public Sub Remove(ByVal Key As IGeneric)
    
    Dim S As Byte
    Dim i As Long: i = FindEntry(Key)
    
    If (i >= 0) Then
        With myClass
            S = (.Buckets(i).State And Not HAS_ENTRY)
            .Buckets(i) = EmptyBucket: .Buckets(i).State = S
            Call InterlockedDecrement(.Count)
        End With
    End If
End Sub

Public Sub Clear()
     Call Skynet.Dispose(Me)
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
    Call Rehash(Expandprime(myClass.HashSize), False)
End Sub

Private Function InitialHashing(ByVal HashValue As Long, ByVal HashSize As Long) As HashItem
    
    With InitialHashing
        .Size = HashSize
        .Index = Modulo(HashValue, .Size)
        .Incr = 1 + Modulo((CDbl(HashValue) * myClass.Const.HashPrime), (.Size - 1))
    End With
    
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

Private Sub PutEntry(ByRef Buckets() As Bucket, ByRef Entry As Bucket, ByRef Hash As HashItem)

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
                    ByVal Value As IGeneric)

    Dim Try As Long
    Dim HashValue As Long
    
    If (Key Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Function Add").KeyIsNothing
        
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

Private Sub Class_Terminate()
    Call Skynet.Dispose(Me)
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
    c = myClass 'hardcopy

    Virtual.Pointer = VarPtr(c)
    Set IGeneric_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function

Private Sub IGeneric_Dispose()
    If myClass.isSealed Then
    
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
                            "HashSize: " & .HashSize & vbNewLine & _
                            "LoadFactor: " & .LoadFactor & vbNewLine & _
                            "LoadSize: " & .LoadSize & vbNewLine & _
                            "Elements: " & .Count & vbNewLine & _
                            "Occupancy: " & .Occupancy
    End With
End Function

'IGenericDictionary
Private Property Get IGenericDictionary_Elements() As IGenericCollection: Set IGenericDictionary_Elements = Me: End Property
Private Property Set IGenericDictionary_Item(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value): End Property
Private Property Get IGenericDictionary_Item(ByVal Key As IGeneric) As IGeneric: Set IGenericDictionary_Item = myClass.Buckets(FindEntry(Key)).Value: End Property
Private Property Get IGenericDictionary_Values() As IGenericCollection
    
    myClass.isSealed = True
    myClass.SealedClass = Values
    IGenericDictionary_Values = Me

End Property
Private Property Get IGenericDictionary_Keys() As IGenericCollection
    
    myClass.isSealed = True
    myClass.SealedClass = Keys
    IGenericDictionary_Keys = Me

End Property
Private Function IGenericDictionary_ContainsKey(ByVal Key As IGeneric) As Boolean: IGenericDictionary_ContainsKey = Me.TryGetValue(Key, Nothing): End Function
Private Function IGenericDictionary_TryGetValue(ByVal Key As IGeneric, Value As IGeneric) As Boolean: IGenericDictionary_TryGetValue = Me.TryGetValue(Key, Value): End Function
Private Sub IGenericDictionary_Remove(ByVal Key As IGeneric): Call Me.Remove(Key): End Sub
Private Sub IGenericDictionary_Add(ByVal Key As IGeneric, ByVal Value As IGeneric): Call Me.Add(Key, Value): End Sub

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Count: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = False: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = myClass.IsReadOnly: End Property
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long)

    If (GArray Is Nothing) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericCollection_CopyTo").NullPointer("GArray Is Nothing")
    
    If (ArrayIndex < GArray.LowerBound) Or (ArrayIndex > GArray.Length) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericCollection_CopyTo").IndexOutofRange("(Index < 0) Or (Index > GArray.Length)")

    If (GArray.Length - ArrayIndex + GArray.LowerBound < Me.Count) Then _
        Call Skynet.Throw(Me, "Private Sub IGenericCollection_CopyTo").InvalidInput("(GArray.Length - Index < Me.Count)")
    
    Select Case myClass.SealedClass
        Case Keys
            Call AddKeysToArray(GArray:=GArray, ArrayIndex:=ArrayIndex)
        Case Values
            Call AddValuesToArray(GArray:=GArray, ArrayIndex:=ArrayIndex)
        Case Else
            Call AddPairsToArray(GArray:=GArray, ArrayIndex:=ArrayIndex)
    End Select
    
End Sub
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Function IGenericCollection_ToArray() As GenericArray
    
    Dim GArray As GenericArray
    Set GArray = GenericArray.Build(myClass.Count)
    
    Call IGenericCollection_CopyTo(GArray:=GArray, ArrayIndex:=GArray.LowerBound)
    Set IGenericCollection_ToArray = GArray
    
End Function
Private Function IGenericCollection_NewEnumerator() As IGenericEnumerator

End Function
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean
    
    Dim i As Long
    
    If myClass.SealedClass = Keys Then
        IGenericCollection_Contains = Me.Contains(Key:=Item)
    Else
        For i = LBound(myClass.Buckets) To UBound(myClass.Buckets)
            If (myClass.Buckets(i).State And HAS_ENTRY) = HAS_ENTRY Then _
                If (Item.Equals(myClass.Buckets(i).Value)) Then _
                    IGenericCollection_Contains = True: Exit Function
            
        Next
    End If
    
End Function

'IPrototype
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericMap: End Function
Private Property Get IPrototype_Default() As Object: Set IPrototype_Default = GenericMap: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property


