VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Skynet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@PredeclaredId
'@Folder "Base"
Option Explicit

Private Const MAXVALUE As Double = 2 ^ 31
Private Const MAX_INT32 As Long = &H7FFFFFFF
Private Const NULLPOINTER As Long = 0

#If Win64 Then
    Private Const POINTERSIZE As LongPtr = 8
    Private Declare PtrSafe Sub CopyMemory_byVar Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDestination As Any, ByRef pSource As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef pDestination As Any, ByVal Length As Long, ByVal Fill As Byte)
    Private Declare PtrSafe Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Size As Long, ByRef Dst As LongPtr, ByVal Src As LongPtr) As Long
    Private Declare PtrSafe Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Size As Long, ByRef Dst As LongPtr, ByVal Src As LongPtr) As Long
    Private Declare PtrSafe Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As LongPtr
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal Size As LongPtr)
#Else

    Private Const POINTERSIZE As Long = 4
    Private Declare Sub CopyMemory_byVar Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDestination As Any, ByRef pSource As Any, ByVal Length As Long)
    Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef pDestination As Any, ByVal Length As Long, ByVal Fill As Byte)
    Private Declare Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Size As Long, ByRef Dst As LongPtr, ByVal Src As Long) As Long
    Private Declare Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Size As Long, ByRef Dst As Long, ByVal Src As Long) As Long
    Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal Size As Long)
#End If
'#If Mac Then
'  #If Win64 Then
'    Private Declare PtrSafe Function CopyMemory_byPtr Lib "libc.dylib" Alias "memmove" (ByVal dest As LongPtr, ByVal Src As LongPtr, ByVal Size As Long) As LongPtr
'    Private Declare PtrSafe Function CopyMemory_byVar Lib "libc.dylib" Alias "memmove" (ByRef dest As Any, ByRef Src As Any, ByVal Size As Long) As LongPtr
'    Private Declare PtrSafe Function ZeroMemory_byVar Lib "libc.dylib" Alias "memmove" (ByRef dest As Any, ByRef Src As Any, ByVal Size As Long) As LongPtr
'  #Else
    'Private Declare Function CopyMemory_byPtr Lib "libc.dylib" Alias "memmove" (ByVal dest As Long, ByVal src As Long, ByVal size As Long) As Long
    'Private Declare Function CopyMemory_byVar Lib "libc.dylib" Alias "memmove" (ByRef dest As Any, ByRef src As Any, ByVal size As Long) As Long
'  #End If

Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long

Public Enum Hashing
    Function1 = 1
    Function2
End Enum

Public Enum SortOrder
    Descending = -1
    Ascending = 1
End Enum

Public Enum SortMethod
    IntrospectiveSort = -1
    DepthLimitedQuickSort = 1
End Enum

Public Enum SystemState
    Ready
    BuildNew
End Enum

Private Type Member
    CallStack As New VBA.Collection
End Type
Private myClass As Member
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents CurrentTask As Task
Attribute CurrentTask.VB_VarHelpID = -1

Public Property Get MaxLong() As Long
    MaxLong = MAX_INT32
End Property

Public Sub Assert(ByVal t As Task)
   Call Err.Raise(t.ErrorNumber, TypeName$(t.Source), Me.Generic(t).ToString)
End Sub

Public Function Contract(ByVal Source As Object, ByVal Method As String) As Task
    Set Contract = Task.Build(Source, Method)
End Function

Public Function Push(ByVal Source As Object, ByVal Method As String) As Task

    Set CurrentTask = Task.Build(Source, Method)
    Call myClass.CallStack.Add(CurrentTask)
    Set Push = CurrentTask
    
End Function

Public Function Peek() As Task
        Set CurrentTask = myClass.CallStack(myClass.CallStack.Count)
        Set Peek = CurrentTask
End Function

Public Function Pop() As Task
    Set Pop = Me.Peek
    
    With myClass.CallStack
        Call .Remove(.Count)
    End With
End Function

Public Property Get PointerLength() As Byte: PointerLength = POINTERSIZE: End Property

Public Sub ClearData(ByVal Dst As LongPtr, ByVal Bytes As Long)
    Call FillMemory(ByVal Dst, Bytes, 0) 'Call ZeroMemory(ByVal dst, Bytes)
End Sub

Public Sub CopyData(ByVal Dst As LongPtr, ByVal Src As LongPtr, ByVal Bytes As Long, Optional ByVal ClearSource As Boolean = False)
    Call CopyMemory_byVar(ByVal Dst, ByVal Src, Bytes) 'Works
    'Call CopyBytes(Bytes, ByVal Dst, ByVal Src) 'works
    If ClearSource Then _
        Call ZeroMemory(ByVal Src, Bytes) 'FillMemory(ByVal Src, Bytes, 0)
End Sub

Public Function Create(ByRef Prototype As IPrototype) As Object
    
    Dim Clone As IPrototype

    With Prototype
        Set Clone = .CreateDefault
        'Call CopyBytes(.Size, ByVal Clone.VirtualPointer, ByVal .VirtualPointer) 'works
        Call CopyMemory_byVar(ByVal Clone.VirtualPointer, ByVal .VirtualPointer, .Size): _
        Call ZeroMemory(ByVal .VirtualPointer, .Size)
    End With

    Set Create = Clone

End Function

Public Function PointerOf(ByRef SomeArray As Variant) As LongPtr
    
    Const VT_BYREF As Long = &H4000
    Dim vType As Long
    
    With Me
        If Not (VarType(SomeArray) And VbVarType.vbArray) = VbVarType.vbArray Then _
            .Push(Me, "ArrayPointer").InvalidInput.Raise ("Not (IsArray(SomeArray))")
    
        Call .CopyData(Dst:=ByVal VarPtr(vType), Src:=ByVal VarPtr(SomeArray), Bytes:=2&)
        Call .CopyData(Dst:=ByVal VarPtr(PointerOf), Src:=ByVal VarPtr(SomeArray) + 8&, Bytes:=.PointerLength)
        'if SomeArray is defined x() as ...
        If (vType And VT_BYREF) = VT_BYREF Then _
            Call .CopyData(Dst:=ByVal VarPtr(PointerOf), Src:=ByVal PointerOf, Bytes:=.PointerLength)
    End With
    
End Function

'SequenceCompare
Public Function SequenceCompare(ByVal a As IGenericList, ByVal b As IGenericList) As ComparerEnum
    
    If (a Is b) Then SequenceCompare = IsEqual: Exit Function
    
    If Not (a Is Nothing) And (b Is Nothing) Then
        
        If Not (Me.Generic(a).IsRelatedTo(b)) Then _
            Me.Push(Me, "CompareLists").InvalidInput.Raise ("Type(a) != Type(b)")
    
        If Not (a.Elements.Count = b.Elements.Count) Then _
            Me.Push(Me, "CompareLists").InvalidInput.Raise ("a.Elements.Count != b.Elements.Count")
        
        Dim i As Long, c As ComparerEnum
        Do While (InterlockedIncrement(i) < a.Elements.Count And c = IsEqual)
            If Not a.Item(i) Is Nothing Then _
                c = a.Item(i).Equals(b.Item(i))
        Loop
    End If
    
    SequenceCompare = c
    
End Function

Public Function GetRefCount(ByRef obj As IUnknown) As Long
    If obj Is Nothing Then _
        Exit Function
    CopyMemory_byVar GetRefCount, ByVal (ObjPtr(obj)) + 4&, 4&
    GetRefCount = GetRefCount - 2
End Function

'https://referencesource.microsoft.com/#mscorlib/system/array.cs,c9d30a83673759f0
Private Function BinarySearch( _
                                ByRef Objects() As IGeneric, _
                                ByRef Value As IGeneric, _
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

Public Function Median(ByVal Low As Long, ByVal hi As Long) As Single
     If (hi - Low) >= 0 Then
        '"Length overflow!"
     End If
     Median = Low + (Me.RightShift(hi - Low, 1))
End Function

'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub Sort(ByRef Objects() As IGeneric, Optional ByVal Order As SortOrder = SortOrder.Ascending)
    Call Quicksort(Objects, LBound(Objects, 1), UBound(Objects, 1), Order)
End Sub

Private Sub Quicksort( _
                        ByRef Elements() As IGeneric, _
                        ByVal lb As Long, _
                        ByVal UB As Long, _
                        ByVal Order As SortOrder)

    Dim i As Long
    Dim j As Long
    
    Dim pivot As IGeneric
    Dim tmp As IGeneric
    
    If (lb > UB) Then Exit Sub
    
    i = lb: j = UB
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
    
    If (j - lb) < (UB - i) Then
        Call Quicksort(Elements, lb, j, Order)
        Call Quicksort(Elements, i, UB, Order)
    Else
        Call Quicksort(Elements, i, UB, Order)
        Call Quicksort(Elements, lb, j, Order)
    End If
    
End Sub

Public Function HashValue( _
                            ByRef Ascii() As Byte, _
                            Optional ByVal Hashfunction As Hashing = Hashing.Function1) As Long
    
    Dim h As Double
    Dim i As Long
    
    Select Case Hashfunction
        Case Hashing.Function1
            
            Const Shift As Long = 2& ^ 5
            Dim h1 As Double: h1 = 352654597# 'LeftShift(5381,16))+5381
            Dim h2 As Double: h2 = h1
            With Me
                For i = LBound(Ascii) To UBound(Ascii) - 1
                    h1 = .X0R((h1 * Shift) + h1, Ascii(i)): h2 = .X0R((h2 * Shift) + h2, Ascii(i + 1))
                Next
            End With
            
            h = h1 + (h2 * 1566083941#)
        
        Case Hashing.Function2
            
             Dim Length As Long: Length = UBound(Ascii) + 1
             Dim N As Long
             
             For N = (Length / 2) To 1 Step -1
                h = h + Ascii(i)
                h = Me.X0R(Me.X0R(LeftShift(h, 16), Me.LeftShift(Ascii(i + 1), 11)), h)
                h = h + Me.RightShift(h, 11)
                i = i + 2
            Next
            
            If ((Length Mod 2) = 1) Then
                h = h + Ascii(i) + 1566083941
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
    If Not (Me Is Skynet) Then _
        Call Me.Push(Me, "Private Sub Class_Initialize()").Instantiation.Raise("System class must not be instantiated")
   
End Sub

Public Function ClassName(ByRef Instance As IGeneric) As String: ClassName = "<" & TypeName$(Instance) & ">": End Function
Public Function Generic(ByRef Instance As IGeneric) As IGeneric: Set Generic = Instance: End Function
Attribute Generic.VB_UserMemId = 0
Public Function Modulo(ByVal a As Double, ByVal M As Double) As Long: Modulo = (a - (Int(a / M) * M)): End Function
Public Function Decrement(ByRef i As Long) As Long: i = (Not -i): Decrement = i: End Function
Public Function Increment(ByRef i As Long) As Long: i = (-(Not i)): Increment = i: End Function
Public Function RightShift(ByVal Value As Long, Shift As Byte) As Double: RightShift = Value / (2& ^ Shift): End Function
Public Function LeftShift(ByVal Value As Long, Shift As Byte) As Double: LeftShift = Value * (2& ^ Shift): End Function
Public Function LimitDouble(ByVal D As Double) As Long: LimitDouble = CLng(D - (Fix(D / MAXVALUE) * MAXVALUE)): End Function
Public Function X0R(ByVal d1 As Double, ByVal d2 As Double) As Long: X0R = CLng(d1 - (Fix(d1 / MAXVALUE) * MAXVALUE)) Xor CLng(d2 - (Fix(d2 / MAXVALUE) * MAXVALUE)): End Function


