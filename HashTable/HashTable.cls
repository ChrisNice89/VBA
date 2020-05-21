VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HashTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MAXVALUE As Double = 2 ^ 31
Private Const MAXLONG As Long = 2 ^ 31 - 1

'Bucketstate
Private Const ZERO As Byte = 0
Private Const IS_EMPTY As Byte = 0
Private Const NOT_IS_EMPTY As Byte = 255
Private Const HAS_COLLISION As Byte = 2 ^ 0
Private Const HAS_ENTRY As Byte = 2 ^ 1

'Iterationstate
Private Const KEY_LOADED As Byte = 2 ^ 1
Private Const VALUE_LOADED As Byte = 2 ^ 2
Private Const HASH_LOADED As Byte = 2 ^ 3

Public Enum Hashing
    Function1 = 1
    Function2
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
    Seed As Long
    Size As Long
    Index As Long
End Type
'
Private Type Bucket
    HashCode As Long
    Key As String
    Value As Variant
    State As Byte
End Type
'
Private Type Iterator
    CurrentEntry As Bucket
    CurrentSlot As Long
    Slots As Long
    Buckets() As Bucket
    IsUpToDate As Boolean
    State As Byte
    Keys() As String
    Values() As Variant
End Type
'
Private Type Member
    Const As Constant
    HashMode As Hashing
    Iterator As Iterator
    Occupancy As Long
    Count As Long
    Capacity As Long
    RawSize As Long
    HashSize As Long
    LoadSize As Long
    LoadFactor As Double
    'HashCodeProvider as IHashCodeProvider
    Buckets() As Bucket
    CurrentSlot As Long
End Type
'
Private myClass As Member

Private Sub Class_Initialize()
    With myClass.Const
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
    Call Me.Build(0, 0.7)
End Sub

Public Function Build( _
                                    ByVal Capacity As Long, _
                                    Optional ByVal LoadFactor As Double = 0.72, _
                                    Optional ByVal HashFunction As Hashing = Hashing.Function1)
    With myClass
        If Capacity < 0 Then
            'error out of range
        End If
        
        If Not (LoadFactor >= 0.1 And LoadFactor <= 1#) Then
            'error out of range
        End If
        
        .LoadFactor = LoadFactor
        .Capacity = Capacity
        .RawSize = .Capacity / .LoadFactor
        .HashMode = HashFunction
        
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
End Function

Public Property Get HashMode() As Hashing
    HashMode = myClass.HashMode
End Property

Public Property Get GetKeys() As String()
    If Me.Snapshot Then _
        GetKeys = myClass.Iterator.Keys
End Property

Public Property Get GetValues() As Variant()
    If Me.Snapshot Then _
        GetValues = myClass.Iterator.Values
End Property

Public Function Snapshot() As Boolean
    With myClass.Iterator
        If Not .IsUpToDate Then _
            .IsUpToDate = Fill(.Keys, .Values)
    
        Snapshot = .IsUpToDate
    End With
End Function

Private Function Fill(ByRef Keys() As String, ByRef Values() As Variant) As Boolean

    Dim i As Long, j As Long
    
    With myClass
        If Not (.Count = 0) Then
            ReDim Keys(1 To .Count)
            ReDim Values(1 To .Count)
    
            For i = LBound(.Buckets, 1) To UBound(.Buckets, 1)
                With .Buckets(i)
                    If (.State And HAS_ENTRY) = HAS_ENTRY Then
                        j = j + 1
                        Keys(j) = .Key
                        If VarType(.Value) = vbObject Then Set Values(j) = .Value Else Values(j) = .Value
                    End If
                End With
            Next
            Fill = True
        End If
    End With

End Function

Public Property Get Count() As Long
    Count = myClass.Count
End Property

Public Property Get CurrentHashValue() As Variant
    With myClass.Iterator
        If (.State And HASH_LOADED) = HASH_LOADED Then
            CurrentHashValue = .CurrentEntry.HashCode: .State = (.State And Not HASH_LOADED)
        Else
            CurrentHashValue = Null
        End If
    End With
End Property

Public Property Get CurrentKey() As Variant
    With myClass.Iterator
          If (.State And KEY_LOADED) = KEY_LOADED Then
            CurrentKey = .CurrentEntry.Key: .State = (.State And Not KEY_LOADED)
        Else
            CurrentKey = Null
        End If
    End With
End Property

Public Property Get CurrentType() As VbVarType: CurrentType = VarType(myClass.Iterator.CurrentEntry.Value): End Property
Public Property Get CurrentItem() As Variant
    With myClass.Iterator
        If (.State And VALUE_LOADED) = VALUE_LOADED Then
            If VarType(.CurrentEntry.Value) = vbObject Then Set CurrentItem = .CurrentEntry.Value Else CurrentItem = .CurrentEntry.Value
        Else
            CurrentItem = Null
        End If
        .State = (.State And Not VALUE_LOADED)
    End With
End Property

Public Sub StartIterator()
     With myClass.Iterator
        .Buckets = myClass.Buckets
        .Slots = myClass.HashSize
        .CurrentSlot = -1
    End With
End Sub

Public Sub ResetIterator()
    With myClass.Iterator
        Erase .Keys
        Erase .Values
        Erase .Buckets
        .IsUpToDate = False
        .State = ZERO
        .CurrentSlot = -1
        .Slots = 0
        .CurrentEntry = EmptyBucket
    End With
End Sub

Public Function EntryLoaded() As Boolean
    
    With myClass.Iterator
        Do While Increment(.CurrentSlot) < .Slots
            If (.Buckets(.CurrentSlot).State And HAS_ENTRY) = HAS_ENTRY Then
                .CurrentEntry = .Buckets(.CurrentSlot)
                .State = KEY_LOADED Or VALUE_LOADED Or HASH_LOADED
                EntryLoaded = True: Exit Function
            End If
        Loop
        .State = ZERO
        .CurrentSlot = -1
    End With
   
End Function

Public Property Get Item(ByVal Key As String) As Variant
    If (FindEntry(Key, myClass.CurrentSlot)) Then
        With myClass.Buckets(myClass.CurrentSlot)
            If VarType(.Value) = vbObject Then Set Item = .Value Else Item = .Value
        End With
    Else
        Item = Null
    End If
End Property

Public Property Get LastAccess() As Variant
    With myClass.Buckets(myClass.CurrentSlot)
        If VarType(.Value) = vbObject Then Set LastAccess = .Value Else LastAccess = .Value
    End With
End Property

Public Function Contains(ByVal Key As String) As Boolean: Contains = FindEntry(Key, myClass.CurrentSlot): End Function
Public Function Add(ByVal Key As String, ByVal Value As Variant): Call Insert(Key, Value): End Function
Public Function Add2(ByVal Key As String, ByVal Value As Variant, ByVal HashValue As Long): Call Insert(Key, Value, HashValue): End Function
Public Property Let Item(ByVal Key As String, ByVal Value As Variant): Call Insert(Key, Value): End Property

Public Function Copy() As HashTable

    Dim i As Long
    Dim ht As HashTable
    Set ht = New HashTable
    
    With myClass
        Call ht.Build(.LoadSize, .LoadFactor, .HashMode)
        
        For i = LBound(.Buckets, 1) To UBound(.Buckets, 1)
            With .Buckets(i)
                If (.State And HAS_ENTRY) = HAS_ENTRY Then _
                    Call ht.Add2(.Key, .Value, .HashCode)
            End With
        Next
    End With
   
    Set Copy = ht

End Function

Public Function RemoveAll()
    With myClass
        Call Me.ResetIterator
        ReDim .Buckets(.HashSize)
        .Count = 0
        .CurrentSlot = -1
    End With
End Function

Public Function RemoveRange(ByRef Keys As Variant)
    Dim K
    If Not IsArray(Keys) Then
        'error invalid type
    End If
    For Each K In Keys
        Call Me.Remove(K)
    Next
End Function

Public Function Remove(ByVal Key As String)
    
    Dim s As Byte
    Dim Slot As Long
    
    With myClass
        If FindEntry(Key, Slot) Then
            .Iterator.IsUpToDate = False
            s = (.Buckets(Slot).State And Not HAS_ENTRY)
            .Buckets(Slot) = EmptyBucket: .Buckets(Slot).State = s
            Call Decrement(.Count)
            
            If (Slot = .CurrentSlot) Then _
                .CurrentSlot = -1
        End If
    End With
    
End Function

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

Public Function HashValue(ByVal Text As String, Optional ByVal HashFunction As Hashing = Hashing.Function1) As Long
    
    Dim h As Double
    Dim i As Long
    Dim ascii() As Byte: ascii() = StrConv(Text, vbFromUnicode)
    
    Select Case HashFunction
        Case Hashing.Function1
            
            Const Shift As Long = 2& ^ 5
            Dim h1 As Double: h1 = 352654597# 'LeftShift(5381,16))+5381
            Dim h2 As Double: h2 = h1
            
            For i = LBound(ascii) To UBound(ascii) - 1
                h1 = X0R((h1 * Shift) + h1, ascii(i)): h2 = X0R((h2 * Shift) + h2, ascii(i + 1))
            Next
            
            h = h1 + (h2 * 1566083941#)
        
        Case Hashing.Function2
            
             Dim length As Long: length = UBound(ascii()) + 1
             Dim n As Long
             
             For n = (length / 2) To 1 Step -1
                h = h + ascii(i)
                h = X0R(X0R(LeftShift(h, 16), LeftShift(ascii(i + 1), 11)), h)
                h = h + RightShift(h, 11)
                i = i + 2
            Next
            
            If ((length Mod 2) = 1) Then
                h = h + ascii(i) + 1566083941
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
Private Function Modulo(ByVal a As Double, ByVal m As Double) As Long: Modulo = (a - (Int(a / m) * m)): End Function
Private Function Decrement(ByRef i As Long) As Long: i = (Not -i): Decrement = i: End Function
Private Function Increment(ByRef i As Long) As Long: i = (-(Not i)): Increment = i: End Function
Private Function RightShift(ByVal Value As Long, Shift As Byte) As Double: RightShift = Value / (2& ^ Shift): End Function
Private Function LeftShift(ByVal Value As Long, Shift As Byte) As Double: LeftShift = Value * (2& ^ Shift): End Function
Private Function LimitDouble(ByVal d As Double) As Long: LimitDouble = CLng(d - (Fix(d / MAXVALUE) * MAXVALUE)): End Function
Private Function X0R(ByVal d1 As Double, ByVal d2 As Double) As Long: X0R = CLng(d1 - (Fix(d1 / MAXVALUE) * MAXVALUE)) Xor CLng(d2 - (Fix(d2 / MAXVALUE) * MAXVALUE)): End Function

Public Function ToString() As String
    With myClass
        ToString = "HashSize: " & .HashSize & vbNewLine & _
                        "LoadFactor: " & .LoadFactor & vbNewLine & _
                        "LoadSize: " & .LoadSize & vbNewLine & _
                        "Elements: " & .Count & vbNewLine & _
                        "Occupancy: " & .Occupancy & vbNewLine & _
                        "HashFunction: " & .HashMode
    End With
End Function

Private Function InitialHashing(ByVal HashValue As Long, ByVal HashSize As Long) As HashItem

    Dim Hash As HashItem
    
    With Hash
        .Seed = HashValue: .Size = HashSize
        .Index = Modulo(.Seed, .Size)
        .Incr = 1 + Modulo((CDbl(.Seed) * myClass.Const.HashPrime), (.Size - 1))
    End With
    
    LSet InitialHashing = Hash

End Function

Private Function Rehash(ByVal Size As Long, ByVal ForceNewHashing As Boolean)

    Dim NewBuckets() As Bucket: ReDim NewBuckets(Size - 1)
    Dim b As Bucket
    Dim i As Long, HasEntry As Boolean
    
    With myClass
        .Occupancy = 0
        
        If ForceNewHashing Then
            For i = LBound(.Buckets) To UBound(.Buckets)
                If (.Buckets(i).State And HAS_ENTRY) = HAS_ENTRY Then _
                    b = .Buckets(i): b.HashCode = Me.HashValue(b.Key, myClass.HashMode): Call PutEntry(NewBuckets, b, InitialHashing(b.HashCode, Size))
            Next
        Else
            For i = LBound(.Buckets) To UBound(.Buckets)
                If (.Buckets(i).State And HAS_ENTRY) = HAS_ENTRY Then _
                    b = .Buckets(i): Call PutEntry(NewBuckets, b, InitialHashing(b.HashCode, Size))
            Next
        End If
        
        'b = EmptyBucket
        
        If .CurrentSlot >= 0 Then _
            b.Key = .Buckets(.CurrentSlot).Key: HasEntry = True

        .Buckets = NewBuckets
        .LoadSize = Int(.LoadFactor * Size)
        .HashSize = Size
         
         If HasEntry Then _
            Call FindEntry(b.Key, .CurrentSlot)
            
    End With
    
End Function

Private Function PutEntry(ByRef Buckets() As Bucket, ByRef Entry As Bucket, ByRef Hash As HashItem)

    Entry.State = (Entry.State And Not HAS_COLLISION)
    
    With Hash
        Do
            If (Buckets(.Index).State Xor IS_EMPTY) = IS_EMPTY Then _
                Buckets(.Index) = Entry: Exit Function
            
             If Not (Buckets(.Index).State And HAS_COLLISION) = HAS_COLLISION Then _
                Buckets(.Index).State = (HAS_ENTRY + HAS_COLLISION): Call Increment(myClass.Occupancy)
                
            .Index = Modulo(.Index + .Incr, .Size)
        Loop
    End With

End Function

Private Function FindEntry(ByVal Key As String, Optional ByRef Slot As Long) As Boolean

    Dim try As Long
    
    With InitialHashing(Me.HashValue(Key, myClass.HashMode), myClass.HashSize)
        Do
            If ((myClass.Buckets(.Index).State And HAS_ENTRY) = HAS_ENTRY) Then
                If (StrComp(myClass.Buckets(.Index).Key, Key, vbBinaryCompare) = ZERO) Then _
                    Slot = .Index: FindEntry = True: Exit Function
            End If
            
            If (Not (myClass.Buckets(.Index).State And HAS_COLLISION) = HAS_COLLISION) Then _
                Slot = -1: FindEntry = False: Exit Function
            
            If ((Increment(try)) > .Size) Then _
                Slot = -1: FindEntry = False: Exit Function
                
            .Index = Modulo(.Index + .Incr, .Size)
        Loop
    End With

End Function

Private Sub Insert(ByVal Key As String, ByRef Value As Variant, Optional ByVal HashValue As Long = 0)
    
    Dim Hash As HashItem
    Dim try As Long
    
    With myClass
        .Iterator.IsUpToDate = False
        If (.Count >= .LoadSize) Then
            Call Rehash(Expandprime(.HashSize), False)
        ElseIf ((.Occupancy > .LoadSize) And (.Count > .Const.RehashBarrier)) Then
            Call Rehash(.HashSize, False)
        End If
        
        If (HashValue = 0) Then
            Hash = InitialHashing(Me.HashValue(Key, .HashMode), .HashSize)
        Else
            Hash = InitialHashing(HashValue, .HashSize)
        End If
    End With
    
    With Hash
        Do
            With myClass.Buckets(.Index)
                If ((.State And HAS_ENTRY) = HAS_ENTRY) Then
                    If (StrComp(.Key, Key, vbBinaryCompare) = ZERO) Then
                        If VarType(Value) = vbObject Then Set .Value = Value Else .Value = Value
                        Exit Sub
                    End If
                Else
                    .State = HAS_ENTRY
                    .HashCode = Hash.Seed
                    .Key = Key
                    If VarType(Value) = vbObject Then Set .Value = Value Else .Value = Value
                    Call Increment(myClass.Count)
                    Exit Sub
                End If
                
                If (Not (.State And HAS_COLLISION) = HAS_COLLISION) Then _
                    .State = (.State Or HAS_COLLISION): Call Increment(myClass.Occupancy)
                
            End With
            
            .Index = Modulo(.Index + .Incr, .Size)
        Loop While (Increment(try) < .Size)
    End With

End Sub



