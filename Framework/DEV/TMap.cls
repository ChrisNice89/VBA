VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@PredeclaredId
'@Folder "<T>Enumerable"
Option Explicit

Implements IObject
Implements IPrototype
Implements IEnumerator

Private Const MAXVALUE As Double = 2 ^ 31
Private Const MAXLONG As Long = 2 ^ 31 - 1

'Bucketstate
Private Const ZERO As Byte = 0
Private Const IS_EMPTY As Byte = 0
Private Const HAS_ENTRY As Byte = 2 ^ 0
Private Const HAS_COLLISION As Byte = 2 ^ 1

Private Enum Data
    Values
    Keys
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
    Key As IObject
    Value As IObject
    State As Byte
End Type

'
Private Type Cache
    CurrentSlot As Long
    Slots As Long
    Buckets() As Bucket
End Type

'
Private Type Member
    Const As Constant
    TKey As IObject
    TValue As IObject
    HashMode As Hashing
    Cache As Cache
    Occupancy As Long
    Count As Long
    Capacity As Long
    RawSize As Long
    HashSize As Long
    LoadSize As Long
    LoadFactor As Double
    Buckets() As Bucket
End Type

'
Private myClass As Member
Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type

Private Virtual As Internal

Public Property Get TValue() As IObject: Set TValue = myClass.TValue: End Property
Public Property Get TKey() As IObject: Set TKey = myClass.TKey: End Property
Public Property Get LoadSize() As Long: LoadSize = myClass.LoadSize: End Property
Public Property Get LoadFactor() As Double: LoadFactor = myClass.LoadFactor: End Property
Public Property Get HashMode() As Hashing: HashMode = myClass.HashMode: End Property
Public Property Get Count() As Long: Count = myClass.Count: End Property
Public Function Contains(ByVal Key As IObject) As Boolean: Contains = TryGetValue(Key, Nothing): End Function
Public Sub Add(ByVal Key As IObject, Optional ByVal Value As IObject): Call Insert(Key, Value): End Sub
Public Property Let Item(ByVal Key As IObject, ByVal Value As IObject): Call Insert(Key, Value): End Property
Public Property Set Item(ByVal Key As IObject, ByVal Value As IObject): Call Insert(Key, Value): End Property

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    
End Sub

'Public Function Build( _
'       ByVal TKey As IObject, _
'       ByVal TValue As IObject, _
'       Optional ByVal Capacity As Long, _
'       Optional ByVal LoadFactor As Double = 0.72, _
'       Optional ByVal Hashfunction As Hashing = Hashing.Function1) As TMap
Public Function Build( _
                        Optional ByVal Capacity As Long, _
                        Optional ByVal LoadFactor As Double = 0.72, _
                        Optional ByVal Hashfunction As Hashing = Hashing.Function1) As TMap
    With myClass
        With .Const
            .Primes = Array(3, 7, 11, 17, 23, 29, 37, 47, 59, 71, 89, 107, 131, 163, 197, 239, 293, 353, 431, 521, 631, 761, 919, _
                            1103, 1327, 1597, 1931, 2333, 2801, 3371, 4049, 4861, 5839, 7013, 8419, 10103, 12143, 14591, _
                            17519, 21023, 25229, 30293, 36353, 43627, 52361, 62851, 75431, 90523, 108631, 130363, 156437, _
                            187751, 225307, 270371, 324449, 389357, 467237, 560689, 672827, 807403, 968897, 1162687, 1395263, _
                            1674319, 2009191, 2411033, 2893249, 3471899, 4166287, 4999559, 5999471, 7199369)
        
            .MaxVal = MAXLONG
            .InitialSize = .Primes(7) 'getMinPrime
            .HashPrime = 101
            .RehashBarrier = 100
        End With
        
        If Not (LoadFactor >= 0.1 And LoadFactor <= 1#) Then
            'error out of range
        End If
        
'        Set .TKey = TValue
'        Set .TValue = TValue
        .LoadFactor = LoadFactor
        .Capacity = Capacity
        .RawSize = .Capacity / .LoadFactor
        .HashMode = Hashfunction
        
        If .RawSize > .Const.InitialSize Then
            .HashSize = getPrime(.RawSize)
        Else
            .HashSize = .Const.InitialSize
        End If
        
        ReDim .Buckets(.HashSize - 1)
        .LoadSize = Int(.LoadFactor * .HashSize)
        
        If (.LoadSize < .HashSize) Then
            'error invalid loadsize
        End If
        
    End With
    
    Set Build = Skynet.Create(Me)
    
End Function

Public Property Get GetKeys() As Variant()
    
    Dim Output() As Variant
    Call Load(Keys, Output)
    GetKeys = Output
    
End Property

Public Property Get GetValues() As Variant()
    
    Dim Output() As Variant
    Call Load(Values, Output)
    GetValues = Output
    
End Property

Public Function TryGetValue(ByVal Key As IObject, ByRef Value As IObject) As Boolean
    
    Dim Try As Long
    
'    If Not myClass.TKey.IsRelatedTo(Key) Then _
'        Skynet.Push(Me, "Public Function TryGetValue").NotEqual(myClass.TKey, Key).Raise
    
    With InitialHashing(Key.HashValue, myClass.HashSize)
        Do
            With myClass.Buckets(.Index)
                If .Key Is Nothing Then _
                    Exit Function
            
                If .Key.Equals(Key) Then _
                    Set Value = .Value: TryGetValue = True: Exit Function
             
                If (Not (.State And HAS_COLLISION) = HAS_COLLISION) Then _
                    Exit Function
            
            End With
            If ((Increment(Try)) > .Size) Then _
                Exit Function
                
            .Index = Modulo(.Index + .Incr, .Size)
        Loop
    End With
  
End Function

Private Sub Load(ByVal Data As Data, ByRef Vector() As Variant)
    
    If myClass.Count = 0 Then _
       Vector = Array(): Exit Sub
    
    Dim i As Long, j As Long
    
    With myClass
        ReDim Vector(1 To .Count)
        Select Case Data
        Case Values
            For i = LBound(.Buckets) To UBound(.Buckets)
                If Not (.Buckets(i).Key Is Nothing) Then _
                    Set Vector(Increment(j)) = .Buckets(i).Value
                
            Next
        Case Keys
            For i = LBound(.Buckets) To UBound(.Buckets)
                 If Not (.Buckets(i).Key Is Nothing) Then _
                    Set Vector(Increment(j)) = .Buckets(i).Key
                
            Next
        Case Else
            'error invalid case
        End Select
    End With
    
End Sub

Public Sub CachePrepare()
    With myClass.Cache
        .Buckets = myClass.Buckets
        .Slots = myClass.HashSize
        .CurrentSlot = -1
    End With
End Sub

Public Sub CacheClear()
    With myClass.Cache
        Erase .Buckets
        .CurrentSlot = -1
        .Slots = 0
    End With
End Sub

Public Function Cached( _
       ByRef Key As IObject, _
       ByRef Value As IObject, _
       Optional ByRef HashValue As Long = -1) As Boolean
    
    With myClass.Cache
        Do While Increment(.CurrentSlot) < .Slots
            If Not (.Buckets(.CurrentSlot).Key Is Nothing) Then
                With .Buckets(.CurrentSlot)
                    Set Key = .Key
                    Set Value = .Value
                End With
                Cached = True: Exit Function
            End If
        Loop
    End With
      
    Call Me.CacheClear

End Function

Public Property Get Item(ByVal Key As IObject) As IObject
    Dim T As IObject
    
    If (Me.TryGetValue(Key, T)) Then _
        Set Item = T
    
End Property

Public Sub RemoveAll()
    With myClass
        ReDim .Buckets(.HashSize)
        .Count = 0
    End With
End Sub

Public Sub RemoveBy(ByRef Keys As Variant)
    Dim K As IObject
    If Not IsArray(Keys) Then
        'error invalid type
    End If
    For Each K In Keys
        Call Me.Remove(K)
    Next
End Sub

Public Sub Remove(ByVal Key As IObject)
    
    Dim s As Byte
    Dim Slot As Long
    
    With myClass
'        If FindBy(Key, Slot) Then
'            s = (.Buckets(Slot).State And Not HAS_ENTRY)
'            .Buckets(Slot) = EmptyBucket: .Buckets(Slot).State = s
'            Call Decrement(.Count)
'
'
'        End If
    End With
    
End Sub

Private Function getPrime(ByVal Min As Long) As Long
    
    Dim p As Long
    
    If Min < 0 Then
        'error
    End If
    
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

Public Function HashValue(ByVal Text As String, Optional ByVal Hashfunction As Hashing = Hashing.Function1) As Long
    
    Dim h As Double
    Dim i As Long
    Dim Ascii() As Byte: Ascii = StrConv(Text, vbFromUnicode)
    
    Select Case Hashfunction
        Case Hashing.Function1
            
            Const Shift As Long = 2& ^ 5
            Dim h1 As Double: h1 = 352654597# 'LeftShift(5381,16))+5381
            Dim h2 As Double: h2 = h1
            
            For i = LBound(Ascii) To UBound(Ascii) - 1
                h1 = X0R((h1 * Shift) + h1, Ascii(i)): h2 = X0R((h2 * Shift) + h2, Ascii(i + 1))
            Next
            
            h = h1 + (h2 * 1566083941#)
        
        Case Hashing.Function2
            
            Dim Length As Long: Length = UBound(Ascii) + 1
            Dim n As Long
             
            For n = (Length / 2) To 1 Step -1
                h = h + Ascii(i)
                h = X0R(X0R(LeftShift(h, 16), LeftShift(Ascii(i + 1), 11)), h)
                h = h + RightShift(h, 11)
                i = i + 2
            Next
            
            If ((Length Mod 2) = 1) Then
                h = h + Ascii(i) + 1566083941
                h = X0R(h, LeftShift(h, 10))
                h = h + RightShift(h, 1)
            End If
            
            h = X0R(h, LeftShift(h, 3)): h = h + RightShift(h, 5)
            h = X0R(h, LeftShift(h, 4)): h = h + RightShift(h, 17)
            h = X0R(h, LeftShift(h, 25)): h = h + RightShift(h, 6)
    End Select
    
    HashValue = CLng(h - (Fix(h / MAXVALUE) * MAXVALUE))

End Function

Private Function Expandprime(ByVal OldSize As Long) As Long
    
    Dim NewSize As Long: NewSize = OldSize * 2
    
    If NewSize > myClass.Const.MaxVal Then
        Expandprime = myClass.Const.MaxVal
    Else
        Expandprime = getPrime(NewSize)
    End If
End Function

Private Sub Expand()
    Call Rehash(Expandprime(myClass.HashSize), False)
End Sub

Private Function getMinPrime() As Long: getMinPrime = myClass.Const.Primes(LBound(myClass.Const.Primes, 1)): End Function
Private Function EmptyBucket() As Bucket: End Function

Private Function Modulo(ByVal A As Double, ByVal M As Double) As Long: Modulo = (A - (Int(A / M) * M)): End Function
Private Function Decrement(ByRef i As Long) As Long: i = (Not -i): Decrement = i: End Function

Private Function Increment(ByRef i As Long) As Long: i = (-(Not i)): Increment = i: End Function
Private Function RightShift(ByVal Value As Long, Shift As Byte) As Double: RightShift = Value / (2& ^ Shift): End Function

Private Function LeftShift(ByVal Value As Long, Shift As Byte) As Double: LeftShift = Value * (2& ^ Shift): End Function
Private Function LimitDouble(ByVal D As Double) As Long: LimitDouble = CLng(D - (Fix(D / MAXVALUE) * MAXVALUE)): End Function

Private Function X0R(ByVal d1 As Double, ByVal d2 As Double) As Long: X0R = CLng(d1 - (Fix(d1 / MAXVALUE) * MAXVALUE)) Xor CLng(d2 - (Fix(d2 / MAXVALUE) * MAXVALUE)): End Function

Private Function InitialHashing(ByVal HashValue As Long, ByVal HashSize As Long) As HashItem

    Dim Hash As HashItem
    
    With Hash
        .Size = HashSize
        .Index = Modulo(HashValue, .Size)
        .Incr = 1 + Modulo((CDbl(HashValue) * myClass.Const.HashPrime), (.Size - 1))
    End With
    
    LSet InitialHashing = Hash

End Function

Private Sub Rehash(ByVal Size As Long, ByVal ForceNewHashing As Boolean)

    Dim NewBuckets() As Bucket: ReDim NewBuckets(Size - 1)
    Dim B As Bucket
    Dim i As Long
  
    With myClass
        .Occupancy = 0
        
        If ForceNewHashing Then
            For i = LBound(.Buckets) To UBound(.Buckets)
                If Not (.Buckets(i).Key Is Nothing) Then _
                    B = .Buckets(i): B.HashCode = B.Key.HashValue: Call PutEntry(NewBuckets, B, InitialHashing(B.HashCode, Size))
            Next
        Else
            For i = LBound(.Buckets) To UBound(.Buckets)
                If Not (.Buckets(i).Key Is Nothing) Then _
                    B = .Buckets(i): Call PutEntry(NewBuckets, B, InitialHashing(B.HashCode, Size))
            Next
        End If
        
        B = EmptyBucket
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
                Buckets(.Index).State = (HAS_ENTRY + HAS_COLLISION): myClass.Occupancy = myClass.Occupancy + 1
                
            .Index = Modulo(.Index + .Incr, .Size)
        Loop
    End With

End Sub

'Private Function FindBy(ByVal Key As IObject, Optional ByRef Slot As Long) As Boolean
'
'    Dim Try As Long
'
'    With InitialHashing(Key.HashValue, myClass.HashSize)
'        Do
'            If Not (myClass.Buckets(.Index).Key Is Nothing) Then
'                If myClass.Buckets(.Index).Key.Equals(Key) Then _
'                    Slot = .Index: FindBy = True: Exit Function
'            End If
'
'            If (Not (myClass.Buckets(.Index).State And HAS_COLLISION) = HAS_COLLISION) Then _
'                Slot = -1: FindBy = False: Exit Function
'
'            If ((Increment(Try)) > .Size) Then _
'                Slot = -1: FindBy = False: Exit Function
'
'            .Index = Modulo(.Index + .Incr, .Size)
'        Loop
'    End With
'
'End Function

Private Sub Insert(ByRef Key As IObject, ByRef Value As IObject)

    Dim Try As Long
    Dim HashValue As Long
    
    With myClass
'        If Not .TKey.IsRelatedTo(Key) Then _
'            Call Skynet.Push(Me, "Private Sub Insert").NotEqual(.TKey, Key).Raise
'
'        If Not .TValue.IsRelatedTo(Value) Then _
'            Call Skynet.Push(Me, "Private Sub Insert").NotEqual(.TValue, Value).Raise
        
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
                    myClass.Count = myClass.Count + 1
                    Exit Sub
                End If
                
                If .Key.Equals(Key) Then _
                    Set .Value = Value: Exit Sub
             
                If (Not (.State And HAS_COLLISION) = HAS_COLLISION) Then _
                    .State = (HAS_ENTRY Or HAS_COLLISION): myClass.Occupancy = myClass.Occupancy + 1
                
            End With
            
            .Index = Modulo(.Index + .Incr, .Size)
        Loop While (Increment(Try) < .Size)
    End With

End Sub

Public Sub FillEnumerator(ByRef Enumerator As IEnumerator)
  
    Dim c As Member
    With c.Cache
        .Buckets = myClass.Buckets
        .Slots = myClass.HashSize
        .CurrentSlot = -1
    End With
  
    Virtual.Pointer = VarPtr(c)
    Set Enumerator = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
    
End Sub

Private Property Get IEnumerator_Current() As Variant

End Property

Private Function IEnumerator_MoveNext() As Boolean
    With myClass.Cache
        Do While Increment(.CurrentSlot) < .Slots
            If Not (.Buckets(.CurrentSlot).Key Is Nothing) Then
                With .Buckets(.CurrentSlot)
'                    Set Key = .Key
'                    Set Value = .Value
                End With
                IEnumerator_MoveNext = True: Exit Function
            End If
        Loop
        
    End With
End Function

Private Sub IEnumerator_Reset()

End Sub

Private Function IObject_Clone() As IObject
    
    Dim c As Member: c = myClass
    Virtual.Pointer = VarPtr(c)
    Set IObject_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function
'TODO TMap IObject_Dispose
Private Sub IObject_Dispose()

End Sub

Private Function IObject_Equals(Other As IObject) As Boolean
    
    If Other Is Nothing Then IObject_Equals = False: Exit Function
    If Not Other.IsRelatedTo(Me) Then IObject_Equals = False: Exit Function
    
    IObject_Equals = (Skynet.CastObject(Me).ToString = Other.ToString)
     
End Function

'TODO TMap IObject_CompareTo
Private Function IObject_CompareTo(Other As IObject) As ComparerEnum
    Dim map As TMap
    If Other.IsRelatedTo(Me) Then
        Set map = Other
        Select Case True
            Case map.Count < Me.Count: IObject_CompareTo = IsGreater
            Case map.Count > Me.Count: IObject_CompareTo = IsLower
            Case Else: IObject_CompareTo = IsEqual
        End Select
    End If
End Function

Private Property Get IObject_HashValue() As Long
    IObject_HashValue = Skynet.HashValue(StrConv(Skynet.CastObject(Me).ToString, vbFromUnicode), myClass.HashMode)
End Property

Private Function IObject_IsRelatedTo(Other As IObject) As Boolean
    IObject_IsRelatedTo = (TypeOf Other Is TMap)
End Function

Private Function IObject_ToString() As String
    With myClass
        IObject_ToString = "TKey: " & TypeName$(.TKey) & vbNewLine & _
                           "TValue: " & TypeName$(.TValue) & vbNewLine & _
                           "HashSize: " & .HashSize & vbNewLine & _
                           "LoadFactor: " & .LoadFactor & vbNewLine & _
                           "LoadSize: " & .LoadSize & vbNewLine & _
                           "Elements: " & .Count & vbNewLine & _
                           "Occupancy: " & .Occupancy & vbNewLine & _
                           "HashFunction: " & .HashMode
    End With
End Function

Private Property Get IPrototype_NewInstance() As IPrototype: Set IPrototype_NewInstance = New TMap: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property

