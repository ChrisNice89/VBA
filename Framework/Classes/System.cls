VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Public Function Create(ByRef Instance As IObject, ByRef Template As IObject, ByVal Size As Long) As IObject
    If Instance.Equals(Template) Then
        'call copybytes(size,Instance.VirtualStruct,Template.VirtualStruct)
        'call ZeroMemory (Template.VirtualStruct,size)
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

Private Sub Class_Initialize()
    If Not (Me Is System) Then
     
        Err.Raise 5, TypeName$(Me), TypeName$(Me) & " (Run-time error '5': Invalid use of New Keyword) - System class must not be instantiated."
        'Call Err.Raise(5, TypeName$(Me), "Invalid instance use " & TypeName$(Me) & ".Build() to instatiate new " & TypeName$(Me) & " objects.")
    End If

End Sub

Public Function IObject(ByRef Instance As IObject) As IObject: Set IObject = Instance: End Function
Public Function Hash(ByRef Instance As IObject) As Long: Hash = Instance.HashValue: End Function
Public Function Modulo(ByVal a As Double, ByVal m As Double) As Long: Modulo = (a - (Int(a / m) * m)): End Function
Public Function Decrement(ByRef i As Long) As Long: i = (Not -i): Decrement = i: End Function
Public Function Increment(ByRef i As Long) As Long: i = (-(Not i)): Increment = i: End Function
Public Function RightShift(ByVal Value As Long, Shift As Byte) As Double: RightShift = Value / (2& ^ Shift): End Function
Public Function LeftShift(ByVal Value As Long, Shift As Byte) As Double: LeftShift = Value * (2& ^ Shift): End Function
Public Function LimitDouble(ByVal d As Double) As Long: LimitDouble = CLng(d - (Fix(d / MAXVALUE) * MAXVALUE)): End Function
Public Function X0R(ByVal d1 As Double, ByVal d2 As Double) As Long: X0R = CLng(d1 - (Fix(d1 / MAXVALUE) * MAXVALUE)) Xor CLng(d2 - (Fix(d2 / MAXVALUE) * MAXVALUE)): End Function

