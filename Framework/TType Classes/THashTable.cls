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
Private Const HAS_ENTRY As Byte = 2 ^ 0
Private Const HAS_COLLISION As Byte = 2 ^ 1

Public Enum Hashing
    Function1 = 1
    Function2
End Enum

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
    Key As String
    Value As Variant
    State As Byte
End Type
'
Private Type Cache
    Prepared As Boolean
    CurrentSlot As Long
    Slots As Long
    Buckets() As Bucket
End Type
'
Private Type Member
    Const As Constant
    HashMode As Hashing
    Cache As Cache
    OCCUPANCY As Long
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
Private myclass As Member

Private Sub Class_Initialize()
    With myclass.Const
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
    With myclass
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

Public Property Get LoadSize() As Long
    LoadSize = myclass.LoadSize
End Property

Public Property Get LoadFactor() As Double
    LoadFactor = myclass.LoadFactor
End Property

Public Property Get HashMode() As Hashing
    HashMode = myclass.HashMode
End Property

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

Private Sub Load(ByVal Data As Data, ByRef Vector() As Variant)
    
    If myclass.Count = 0 Then _
        Vector = Array(): Exit Sub
    
    Dim i As Long, j As Long
    
    With myclass
        ReDim Vector(1 To .Count)
        Select Case Data
            Case Values
                For i = LBound(.Buckets) To UBound(.Buckets)
                    If .Buckets(i).State And HAS_ENTRY = HAS_ENTRY Then _
                        If VarType(.Buckets(i).Value) = vbObject Then _
                            Set Vector(Increment(j)) = .Buckets(i).Value _
                        Else Vector(Increment(j)) = .Buckets(i).Value
                
                Next
            Case Keys
                 For i = LBound(.Buckets) To UBound(.Buckets)
                    If .Buckets(i).State And HAS_ENTRY = HAS_ENTRY Then _
                        Vector(Increment(j)) = .Buckets(i).Key
                
                Next
            Case Else
                'error invalid case
        End Select
    End With
    
End Sub

Public Property Get Count() As Long
    Count = myclass.Count
End Property

Public Sub CachePrepare()
     With myclass.Cache
        .Prepared = True
        .Buckets = myclass.Buckets
        .Slots = myclass.HashSize
        .CurrentSlot = -1
    End With
End Sub

Public Sub CacheClear()
    With myclass.Cache
        .Prepared = False
        Erase .Buckets
        .CurrentSlot = -1
        .Slots = 0
    End With
End Sub

Public Function Cached( _
                                        ByRef Key As String, _
                                        ByRef Value As Variant, _
                                        Optional ByRef HashValue As Long = -1) As Boolean
    
    With myclass.Cache
        Do While Increment(.CurrentSlot) < .Slots
            If (.Buckets(.CurrentSlot).State And HAS_ENTRY) = HAS_ENTRY Then
                With .Buckets(.CurrentSlot)
                    Key = .Key
                    If VarType(.Value) = vbObject Then _
                        Set Value = .Value _
                    Else Value = .Value
                    
                    If Not HashValue = -1 Then _
                        HashValue = .HashCode
                
                End With
                Cached = True: Exit Function
            End If
        Loop
        
        If Not .Prepared Then
            'error
        End If
    End With
      
    Call Me.CacheClear

End Function

Public Property Get Item(ByVal Key As String) As Variant
    If (FindBy(Key, myclass.CurrentSlot)) Then
        With myclass.Buckets(myclass.CurrentSlot)
            If VarType(.Value) = vbObject Then _
                Set Item = .Value _
            Else Item = .Value
        
        End With
    Else
        Item = Null
    End If
End Property

Public Property Get LastCheck() As Variant
    With myclass.Buckets(myclass.CurrentSlot)
        If VarType(.Value) = vbObject Then _
            Set LastCheck = .Value _
        Else LastCheck = .Value
    End With
End Property

Public Function Contains(ByVal Key As String) As Boolean: Contains = FindBy(Key, myclass.CurrentSlot): End Function
Public Function Add(ByVal Key As String, ByVal Value As Variant): Call Insert(Key, Value): End Function
Public Property Let Item(ByVal Key As String, ByVal Value As Variant): Call Insert(Key, Value): End Property

Public Sub CloneBy(ByVal Other As HashTable)
    
    Dim k As String
    Dim v As Variant
    Dim h As Long
  
    With Other
        Call Me.Build(.LoadSize, .LoadFactor, .HashMode)
        Call .CachePrepare
        Do While .Cached(k, v, h): Call Insert(k, v, h): Loop
    End With
    
End Sub

Public Function RemoveAll()
    With myclass
        ReDim .Buckets(.HashSize)
        .Count = 0
        .CurrentSlot = -1
    End With
End Function

Public Function RemoveBy(ByRef Keys As Variant)
    Dim k
    If Not IsArray(Keys) Then
        'error invalid type
    End If
    For Each k In Keys
        Call Me.Remove(k)
    Next
End Function

Public Function Remove(ByVal Key As String)
    
    Dim s As Byte
    Dim Slot As Long
    
    With myclass
        If FindBy(Key, Slot) Then
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
    
    With myclass.Const
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
            
             Dim Length As Long: Length = UBound(ascii()) + 1
             Dim n As Long
             
             For n = (Length / 2) To 1 Step -1
                h = h + ascii(i)
                h = X0R(X0R(LeftShift(h, 16), LeftShift(ascii(i + 1), 11)), h)
                h = h + RightShift(h, 11)
                i = i + 2
            Next
            
            If ((Length Mod 2) = 1) Then
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
    
    If NewSize > myclass.Const.MaxVal Then
        Expandprime = myclass.Const.MaxVal
    Else
        Expandprime = getPrime(NewSize)
    End If
End Function

Private Sub Expand()
    Call Rehash(Expandprime(myclass.HashSize), False)
End Sub

Private Function getMinPrime() As Long: getMinPrime = myclass.Const.Primes(LBound(myclass.Const.Primes, 1)): End Function
Private Function EmptyBucket() As Bucket: End Function
Private Function Modulo(ByVal a As Double, ByVal m As Double) As Long: Modulo = (a - (Int(a / m) * m)): End Function
Private Function Decrement(ByRef i As Long) As Long: i = (Not -i): Decrement = i: End Function
Private Function Increment(ByRef i As Long) As Long: i = (-(Not i)): Increment = i: End Function
Private Function RightShift(ByVal Value As Long, Shift As Byte) As Double: RightShift = Value / (2& ^ Shift): End Function
Private Function LeftShift(ByVal Value As Long, Shift As Byte) As Double: LeftShift = Value * (2& ^ Shift): End Function
Private Function LimitDouble(ByVal d As Double) As Long: LimitDouble = CLng(d - (Fix(d / MAXVALUE) * MAXVALUE)): End Function
Private Function X0R(ByVal d1 As Double, ByVal d2 As Double) As Long: X0R = CLng(d1 - (Fix(d1 / MAXVALUE) * MAXVALUE)) Xor CLng(d2 - (Fix(d2 / MAXVALUE) * MAXVALUE)): End Function

Public Function ToString() As String
    With myclass
        ToString = "HashSize: " & .HashSize & vbNewLine & _
                        "LoadFactor: " & .LoadFactor & vbNewLine & _
                        "LoadSize: " & .LoadSize & vbNewLine & _
                        "Elements: " & .Count & vbNewLine & _
                        "Occupancy: " & .OCCUPANCY & vbNewLine & _
                        "HashFunction: " & .HashMode
    End With
End Function

Private Function InitialHashing(ByVal HashValue As Long, ByVal HashSize As Long) As HashItem

    Dim Hash As HashItem
    
    With Hash
        '.Seed = HashValue:
        .Size = HashSize
        .Index = Modulo(HashValue, .Size)
        .Incr = 1 + Modulo((CDbl(HashValue) * myclass.Const.HashPrime), (.Size - 1))
    End With
    
    LSet InitialHashing = Hash

End Function

Private Function Rehash(ByVal Size As Long, ByVal ForceNewHashing As Boolean)

    Dim NewBuckets() As Bucket: ReDim NewBuckets(Size - 1)
    Dim b As Bucket
    Dim i As Long, HasEntry As Boolean
    
    With myclass
        .OCCUPANCY = 0
        
        If ForceNewHashing Then
            For i = LBound(.Buckets) To UBound(.Buckets)
                If (.Buckets(i).State And HAS_ENTRY) = HAS_ENTRY Then _
                    b = .Buckets(i): b.HashCode = Me.HashValue(b.Key, myclass.HashMode): Call PutEntry(NewBuckets, b, InitialHashing(b.HashCode, Size))
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
            Call FindBy(b.Key, .CurrentSlot)
            
    End With
    
End Function

Private Function PutEntry(ByRef Buckets() As Bucket, ByRef Entry As Bucket, ByRef Hash As HashItem)

    Entry.State = HAS_ENTRY
    
    With Hash
        Do
            If (Buckets(.Index).State = IS_EMPTY) Then _
                Buckets(.Index) = Entry: Exit Function
            
             If Not (Buckets(.Index).State And HAS_COLLISION) = HAS_COLLISION Then _
                Buckets(.Index).State = (HAS_ENTRY + HAS_COLLISION): Call Increment(myclass.OCCUPANCY)
                
            .Index = Modulo(.Index + .Incr, .Size)
        Loop
    End With

End Function

Private Function FindBy(ByVal Key As String, Optional ByRef Slot As Long) As Boolean

    Dim Try As Long
    
    With InitialHashing(Me.HashValue(Key, myclass.HashMode), myclass.HashSize)
        Do
            If ((myclass.Buckets(.Index).State And HAS_ENTRY) = HAS_ENTRY) Then
                If (StrComp(myclass.Buckets(.Index).Key, Key, vbBinaryCompare) = ZERO) Then _
                    Slot = .Index: FindBy = True: Exit Function
            End If
            
            If (Not (myclass.Buckets(.Index).State And HAS_COLLISION) = HAS_COLLISION) Then _
                Slot = -1: FindBy = False: Exit Function
            
            If ((Increment(Try)) > .Size) Then _
                Slot = -1: FindBy = False: Exit Function
                
            .Index = Modulo(.Index + .Incr, .Size)
        Loop
    End With

End Function

Private Sub Insert(ByRef Key As String, ByRef Value As Variant, Optional ByRef HashValue As Long = -1)

    Dim Try As Long
    
    With myclass
        If (.Count >= .LoadSize) Then
            Call Rehash(Expandprime(.HashSize), False)
        ElseIf ((.OCCUPANCY > .LoadSize) And (.Count > .Const.RehashBarrier)) Then
            Call Rehash(.HashSize, False)
        End If
        
        If (HashValue = -1) Then _
            HashValue = Me.HashValue(Key, .HashMode)
      
    End With
    
    With InitialHashing(HashValue, myclass.HashSize)
        Do
            With myclass.Buckets(.Index)
                If ((.State And HAS_ENTRY) = HAS_ENTRY) Then
                    If (StrComp(.Key, Key, vbBinaryCompare) = ZERO) Then
                        If VarType(Value) = vbObject Then _
                            Set .Value = Value _
                        Else .Value = Value
                        Exit Sub
                    End If
                Else
                    .State = HAS_ENTRY
                    .HashCode = HashValue
                    .Key = Key
                    If VarType(Value) = vbObject Then _
                        Set .Value = Value _
                    Else .Value = Value
                    Call Increment(myclass.Count)
                    Exit Sub
                End If
                
                If (Not (.State And HAS_COLLISION) = HAS_COLLISION) Then _
                    .State = (HAS_ENTRY Or HAS_COLLISION): Call Increment(myclass.OCCUPANCY)
                
            End With
            
            .Index = Modulo(.Index + .Incr, .Size)
        Loop While (Increment(Try) < .Size)
    End With

End Sub



