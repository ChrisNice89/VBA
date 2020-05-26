VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

#If Mac Then
  #If Win64 Then
    Private Declare PtrSafe Function CopyMemory_byPtr Lib "libc.dylib" Alias "memmove" (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal Size As Long) As LongPtr
    Private Declare PtrSafe Function CopyMemory_byVar Lib "libc.dylib" Alias "memmove" (ByRef dest As Any, ByRef src As Any, ByVal Size As Long) As LongPtr
    Private Declare PtrSafe Function ZeroMemory_byVar Lib "libc.dylib" Alias "memmove" (ByRef dest As Any, ByRef src As Any, ByVal Size As Long) As LongPtr
  #Else
    'Private Declare Function CopyMemory_byPtr Lib "libc.dylib" Alias "memmove" (ByVal dest As Long, ByVal src As Long, ByVal size As Long) As Long
    'Private Declare Function CopyMemory_byVar Lib "libc.dylib" Alias "memmove" (ByRef dest As Any, ByRef src As Any, ByVal size As Long) As Long
  #End If
#ElseIf VBA7 Then
  #If Win64 Then
    Private Declare PtrSafe Sub CopyMemory_byPtr Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal Size As LongLong)
    Private Declare PtrSafe Sub CopyMemory_byVar Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal Size As LongLong)
  #Else
    Private Declare PtrSafe Sub CopyMemory_byPtr Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal Size As Long)
    Private Declare PtrSafe Sub CopyMemory_byVar Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal Size As Long)
  #End If
#Else
  'Private Declare Sub CopyMemory_byPtr Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As Long, ByVal src As Long, ByVal size As Long)
  'Private Declare Sub CopyMemory_byVar Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal size As Long)
#End If

''#If VBA7 Then
'    Private Declare PtrSafe Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal ByteLen As Long, ByRef Destination As LongPtr, ByVal Source As LongPtr) As Long
'    Private Declare PtrSafe Sub ZeroMemory Lib "Kernel32" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal cbCopy As Long)
'    Private Declare PtrSafe Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByRef Dst As Any, ByRef Scr As Any, ByVal cbCopy As Long)
''#Else
''    Private Declare Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal ByteLen As Long, ByRef Destination As Long, ByVal Source As Long) As Long
''    Private Declare Sub ZeroMemory Lib "Kernel32" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal cbCopy As Long)
''    Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByRef Dst As Any, ByRef Scr As Any, ByVal cbCopy As Long)
''#End If

Const MAXVALUE As Double = 2 ^ 31
Const MAXLONG As Long = 2 ^ 31 - 1

Public Enum Hashing
    Function1 = 1
    Function2
End Enum

Public Enum SortOrder
    Descending = -1
    Ascending = 1
End Enum

Public Enum ComparerEnum
    IsLower = -1
    IsEqual = 0
    IsGreater = 1
End Enum

Public Enum SystemState
    Ready
    BuildNew
End Enum
Private s As SystemState

Public Property Get State() As SystemState
        State = s
End Property

Public Function ErrorRaised(ByVal Source As IObject, ByVal Message As String) As Error

    Set ErrorRaised = Me.Create(New Error, Error, 12)
End Function

Public Sub CopyData(ByVal Dst As LongPtr, ByVal src As LongPtr, ByVal Bytes As Long)
    Call CopyMemory_byVar(ByVal Dst, ByVal src, Bytes) 'Works
    'Call CopyBytes(Bytes, ByVal Dst, ByVal Src) 'works
End Sub

Public Function Create(ByVal Instance As IObject, ByVal Template As IObject, ByVal Size As Long) As IObject
    Dim z As Long: z = 0
    If Instance.Equals(Template) Then
        'Call CopyBytes(Size, ByVal Instance.VirtualPointer, ByVal Template.VirtualPointer)'works
        'CopyMemory_byVar ByVal Instance.VirtualPointer, ByVal Template.VirtualPointer, Size
         CopyMemory_byVar Instance.VirtualPointer, Template.VirtualPointer, Size
        'Call ZeroMemory(ByVal Template.VirtualPointer, Size) ' works on win
        Set Create = Instance
    End If
End Function

Public Function Search( _
                                            ByRef Objects() As IObject, _
                                            ByRef Value As IObject, _
                                            Optional ByRef IsSorted As Boolean = True) As Long

    If Not IsSorted Then _
        Call Me.Sort(Objects): IsSorted = True

    Search = BinarySearch(Objects, Value, LBound(Objects, 1), UBound(Objects, 1))
    
End Function
'https://referencesource.microsoft.com/#mscorlib/system/array.cs,c9d30a83673759f0
Private Function BinarySearch( _
                                                 ByRef Objects() As IObject, _
                                                 ByRef Value As IObject, _
                                                 ByVal Index As Long, _
                                                 ByVal Length As Long) As Long
                                                

    Dim i As Long
    Dim c As ComparerEnum
    Dim lo As Long: lo = Index
    Dim hi As Long: hi = Index + Length - 1
    
    Do While (lo <= hi)
        i = Me.Median(lo, hi)
        c = Objects(i).CompareTo(Value)
        Select Case c
            Case IsGreater: hi = i - 1
            Case IsLower: lo = i + 1
            Case Else: BinarySearch = i: Exit Function
        End Select
    
    Loop
    
    BinarySearch = Not lo
    
End Function

Public Function Median(ByVal low As Long, ByVal hi As Long) As Single
     If (hi - low) >= 0 Then
        '"Length overflow!"
     End If
     Median = low + (Me.RightShift(hi - low, 1))
End Function

Public Sub Sort(ByRef Objects() As IObject, Optional ByVal Order As SortOrder = SortOrder.Ascending)
    Call Quicksort(Objects, LBound(Objects, 1), UBound(Objects, 1), Order)
End Sub

Private Sub Quicksort(ByRef Elements() As IObject, ByVal LB As Long, ByVal UB As Long, ByVal Order As SortOrder)

    Dim i As Long
    Dim j As Long
    
    Dim pivot As IObject
    Dim tmp As IObject
    
    If (LB > UB) Then Exit Sub
    
    i = LB: j = UB
    Set pivot = Elements((i + j) / 2)
    
    Do
        Do While (Elements(i).CompareTo(pivot) = IsLower * Order): i = i + 1: Loop
        Do While (Elements(j).CompareTo(pivot) = IsGreater * Order): j = j + 1: Loop
        
        If (i <= j) Then
            Set tmp = Elements(i)
            Set Elements(i) = Elements(j)
            Set Elements(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop Until (i > j)
    
    If (j - LB) < (UB - i) Then
        Call Quicksort(Elements, LB, j, Order)
        Call Quicksort(Elements, i, UB, Order)
    Else
        Call Quicksort(Elements, i, UB, Order)
        Call Quicksort(Elements, LB, j, Order)
    End If
    
End Sub

Public Function HashValue(ByRef ascii() As Byte, Optional ByVal HashFunction As Hashing = Hashing.Function1) As Long
    
    Dim h As Double
    Dim i As Long
    
    Select Case HashFunction
        Case Hashing.Function1
            
            Const Shift As Long = 2& ^ 5
            Dim h1 As Double: h1 = 352654597# 'LeftShift(5381,16))+5381
            Dim h2 As Double: h2 = h1
            
            For i = LBound(ascii) To UBound(ascii) - 1
                h1 = Me.X0R((h1 * Shift) + h1, ascii(i)): h2 = Me.X0R((h2 * Shift) + h2, ascii(i + 1))
            Next
            
            h = h1 + (h2 * 1566083941#)
        
        Case Hashing.Function2
            
             Dim Length As Long: Length = UBound(ascii()) + 1
             Dim n As Long
             
             For n = (Length / 2) To 1 Step -1
                h = h + ascii(i)
                h = Me.X0R(Me.X0R(LeftShift(h, 16), Me.LeftShift(ascii(i + 1), 11)), h)
                h = h + Me.RightShift(h, 11)
                i = i + 2
            Next
            
            If ((Length Mod 2) = 1) Then
                h = h + ascii(i) + 1566083941
                h = Me.X0R(h, LeftShift(h, 10))
                h = h + Me.RightShift(h, 1)
            End If
            
            h = Me.X0R(h, Me.LeftShift(h, 3)): h = h + Me.RightShift(h, 5)
            h = Me.X0R(h, Me.LeftShift(h, 4)): h = h + Me.RightShift(h, 17)
            h = Me.X0R(h, Me.LeftShift(h, 25)): h = h + Me.RightShift(h, 6)
    End Select
    
    HashValue = CLng(h - (Fix(h / MAXVALUE) * MAXVALUE))

End Function





Private Sub Class_Initialize()
    If Not (Me Is System) Then
     
        Err.Raise 5, TypeName$(Me), TypeName$(Me) & " (Run-time error '5': Invalid use of New Keyword) - System class must not be instantiated."
        'Call Err.Raise(5, TypeName$(Me), "Invalid instance use " & TypeName$(Me) & ".Build() to instatiate new " & TypeName$(Me) & " objects.")
    End If

End Sub

Public Function IObject(ByRef Instance As IObject) As IObject: Set IObject = Instance: End Function
Public Function Modulo(ByVal a As Double, ByVal m As Double) As Long: Modulo = (a - (Int(a / m) * m)): End Function
Public Function Decrement(ByRef i As Long) As Long: i = (Not -i): Decrement = i: End Function
Public Function Increment(ByRef i As Long) As Long: i = (-(Not i)): Increment = i: End Function
Public Function RightShift(ByVal Value As Long, Shift As Byte) As Double: RightShift = Value / (2& ^ Shift): End Function
Public Function LeftShift(ByVal Value As Long, Shift As Byte) As Double: LeftShift = Value * (2& ^ Shift): End Function
Public Function LimitDouble(ByVal d As Double) As Long: LimitDouble = CLng(d - (Fix(d / MAXVALUE) * MAXVALUE)): End Function
Public Function X0R(ByVal d1 As Double, ByVal d2 As Double) As Long: X0R = CLng(d1 - (Fix(d1 / MAXVALUE) * MAXVALUE)) Xor CLng(d2 - (Fix(d2 / MAXVALUE) * MAXVALUE)): End Function

