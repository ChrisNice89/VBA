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

Implements IGenericIterator

Public Event Dispose()
Public Event Delegate(ByRef Object As IGeneric)

Private Const DEFAULT_SORTORDER As Integer = 1 'SortOrder.Ascending
Private Const MAXVALUE As Double = 2 ^ 31
Private Const MAX_INT32 As Long = &H7FFFFFFF
Private Const NullPointer As Long = 0

#If Win64 Then
    Private Const POINTERSIZE As LongPtr = 8
    Private Declare PtrSafe Sub CopyMemory_byVar Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDestination As Any, ByRef pSource As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef pDestination As Any, ByVal Length As Long, ByVal Fill As Byte)
    Private Declare PtrSafe Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Size As Long, ByRef dst As LongPtr, ByVal src As LongPtr) As Long
    Private Declare PtrSafe Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Size As Long, ByRef dst As LongPtr, ByVal src As LongPtr) As Long
    Private Declare PtrSafe Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As LongPtr
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dst As Any, ByVal Size As LongPtr)
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
    Private Const POINTERSIZE As Long = 4
    Private Declare Sub CopyMemory_byVar Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDestination As Any, ByRef pSource As Any, ByVal Length As Long)
    Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef pDestination As Any, ByVal Length As Long, ByVal Fill As Byte)
    Private Declare Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Size As Long, ByRef dst As Long, ByVal src As Long) As Long
    Private Declare Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Size As Long, ByRef dst As Long, ByVal src As Long) As Long
    Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dst As Any, ByVal Size As Long)
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If
'#If Mac Then
'  #If Win64 Then
'    Private Declare PtrSafe Function CopyMemory_byPtr Lib "libc.dylib" Alias "memmove" (ByVal dest As LongPtr, ByVal Src As LongPtr, ByVal Size As Long) As LongPtr
'    Private Declare PtrSafe Function CopyMemory_byVar Lib "libc.dylib" Alias "memmove" (ByRef dest As Any, ByRef Src As Any, ByVal Size As Long) As LongPtr
'    Private Declare PtrSafe Function ZeroMemory_byVar Lib "libc.dylib" Alias "memmove" (ByRef dest As Any, ByRef Src As Any, ByVal Size As Long) As LongPtr
'  #Else
    'Private Declare Function CopyMemory_byPtr Lib "libc.dylib" Alias "memmove" (ByVal dest As Long, ByVal src As Long, ByVal Size As Long) As Long
    'Private Declare Function CopyMemory_byVar Lib "libc.dylib" Alias "memmove" (ByRef dest As Any, ByRef src As Any, ByVal Size As Long) As Long
'  #End If

Public Enum Hashing
    Function1 = 1
    Function2
End Enum

Public Enum SortMethod
    IntrospectiveSort = -1
    DepthLimitedQuicksort = 1
End Enum

Public Enum SystemState
    Ready
    BuildNew
End Enum

Private Type Range
    Start As Long
    Last As Long
    i As Long
    Current As IGeneric
End Type

Private Type Member
     Range As Range
End Type
Private myclass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get SortOrder() As SortOrder: SortOrder = DEFAULT_SORTORDER: End Property
Public Property Get PointerLength() As Byte: PointerLength = POINTERSIZE: End Property
Public Property Get MAXLONG() As Long: MAXLONG = MAX_INT32: End Property
Public Property Get MaxHash() As Double: MaxHash = MAXVALUE: End Property

Public Property Get Range(ByVal Start As Long, ByVal Last As Long) As IGenericIterator
    
    Dim c As Member
    
    If Start > Last Then _
        Call Me.Throw(Me, "Public Function Range").InvalidInput("Start > Last")
        
    With c.Range
        .Start = Start
        .Last = Last
        .i = .Start - 1
    End With
   
    Set Range = Skynet.CreateInstance(New Skynet, VarPtr(c), LenB(c))

End Property

Public Function Throw(ByVal Source As Object, ByVal Method As String) As Object
    Set Throw = GenericThrowhelper.Build(Source, Method)
End Function

Public Sub ClearData(ByVal dst As LongPtr, ByVal Bytes As Long)
    Call FillMemory(ByVal dst, Bytes, 0) 'Call ZeroMemory(ByVal dst, Bytes)
End Sub

Public Sub CopyData(ByVal dst As LongPtr, ByVal src As LongPtr, ByVal Bytes As Long, Optional ByVal ClearSource As Boolean = False)
    
    Call CopyMemory_byVar(ByVal dst, ByVal src, Bytes) 'Works
    'Call CopyBytes(Bytes, ByVal dst, ByVal src) 'works
    If ClearSource Then _
        Call ZeroMemory(ByVal src, Bytes) 'FillMemory(ByVal Src, Bytes, 0)
End Sub

Public Function CreateInstance(ByVal Prototype As IGeneric, ByVal Source As LongPtr, ByVal Size As Long) As Object
        
    With Prototype
'        If (Prototype Is .StaticInstance) Then _
'            Set Prototype = .CreateDefault
        'Call CopyBytes(.Size, ByVal Clone.VirtualPointer, ByVal .VirtualPointer) 'works
        Call CopyMemory_byVar(ByVal .VirtualPointer, ByVal Source, Size): Call ZeroMemory(ByVal Source, Size)
    End With

    Set CreateInstance = Prototype

End Function

Public Function CreateInstanceOf(ByVal Prototype As IGeneric) As Object
    
    Dim Instance As IGeneric
    
    With Prototype
        Set Instance = .Default
        'Call CopyBytes(.Size, ByVal Clone.VirtualPointer, ByVal .VirtualPointer) 'works
        Call CopyMemory_byVar(ByVal Instance.VirtualPointer, ByVal .VirtualPointer, .VirtualSize): Call ZeroMemory(ByVal .VirtualPointer, .VirtualSize)
    End With

    Set CreateInstanceOf = Prototype

End Function

Public Function PointerOf(ByRef SomeArray As Variant) As LongPtr

    Const VT_BYREF As Long = &H4000
    Dim vType As Long

    With Me
        If Not (VarType(SomeArray) And VbVarType.vbArray) = VbVarType.vbArray Then _
            .Throw(Me, "ArrayPointer").InvalidInput ("Not (IsArray(SomeArray))")

        Call .CopyData(dst:=ByVal VarPtr(vType), src:=ByVal VarPtr(SomeArray), Bytes:=2&)
        Call .CopyData(dst:=ByVal VarPtr(PointerOf), src:=ByVal VarPtr(SomeArray) + 8&, Bytes:=POINTERSIZE)

        'if SomeArray is defined x() as ...
        If (vType And VT_BYREF) = VT_BYREF Then _
            Call .CopyData(dst:=ByVal VarPtr(PointerOf), src:=ByVal PointerOf, Bytes:=POINTERSIZE)
    End With

End Function

'SequenceCompare
Public Function SequenceCompare(ByVal a As IGenericReadOnlyList, ByVal b As IGenericReadOnlyList) As ComparerEnum
    
    If (a Is b) Then SequenceCompare = IsEqual: Exit Function
    
    If Not (a Is Nothing) And (b Is Nothing) Then
        
        If Not (Me.Generic(a).IsRelatedTo(b)) Then _
            Me.Throw(Me, "CompareLists").InvalidInput ("Type(a) != Type(b)")
    
        If Not (a.Count = b.Count) Then _
            Me.Throw(Me, "CompareLists").InvalidInput ("a.Count != b.Count")
        
        Dim i As Long, c As ComparerEnum
        Do While (InterlockedIncrement(i) < a.Count And c = IsEqual)
            If Not a(i) Is Nothing Then _
                c = a(i).Equals(b(i))
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


'Searches a section of the list for a given element using a binary search
'algorithm. Elements of the list are compared to the search value using the IGeneric
'comparer, which in that case must be implemented by all elements of the
'list and the given search value. This method assumes that the given
'section of the list is already sorted; if this is not the case, the
'result will be incorrect.
'
'The method returns the index of the given value in the list. If the
'list does not contain the given value, the method returns a negative
'integer. The bitwise complement operator (~) can be applied to a
'negative result to produce the index of the first element (if any) that
'is larger than the given search value. This is also the index at which
'the search value should be inserted into the list in order for the list
'to remain sorted.
Public Function BinarySearch( _
                                ByRef Objects() As IGenericValue, _
                                ByVal Value As IGeneric, _
                                ByVal index As Long, _
                                ByVal Length As Long, _
                                ByVal Order As SortOrder) As Long

    Dim i As Long
    Dim c As ComparerEnum
    Dim lo As Long: lo = index
    Dim hi As Long: hi = index + Length - 1
    
    Do While (lo <= hi)
    
        i = lo + ((hi - lo) / 2)
      
        If (Objects(i) Is Nothing) = False Then
            c = Objects(i).CompareTo(Value) * Order
        Else
            c = IIf(Value Is Nothing, ComparerEnum.IsEqual, ComparerEnum.IsLower * Order)
        End If
        
        If (c = IsEqual) Then _
            BinarySearch = i: Exit Function
        
        If (c = IsLower) Then
            lo = i + 1
        Else
            hi = i - 1
        End If

    Loop
    
    BinarySearch = Not lo
    
End Function

'Returns the index of the first occurrence of a given value in a range of
'this list. It does a linear, O(n) search. The list is searched forwards, starting at index
'and upto count number of elements. The
'elements of the list are compared to the given value using the
'Object.Equals method.

Public Function IndexOf( _
                            ByRef Objects() As IGeneric, _
                            ByVal Value As IGeneric, _
                            ByVal StartIndex As Long, _
                            ByVal Count As Long) As Long
    Dim i As Long
    Dim EndIndex As Long

    EndIndex = StartIndex + Count - LBound(Objects)
    
    If Not (Value Is Nothing) Then
        For i = StartIndex To EndIndex
            If Value.Equals(Objects(i)) Then _
                IndexOf = i: Exit Function
        Next
    Else
        For i = StartIndex To EndIndex
            If (Objects(i) Is Nothing) Then _
                IndexOf = i: Exit Function
        Next
    End If
    
    IndexOf = -1
    
End Function
'Reverses the elements in a range of this list. Following a call to this
'method, an element in the range given by index and count
'which was previously located at index i will now be located at
'index index + (index + count - i - 1).
Public Sub Reverse( _
                        ByRef Objects() As IGeneric, _
                        ByVal index As Long, _
                        ByVal Length As Long)

    Dim i As Long: i = index
    Dim j As Long: j = index + Length - 1
    Dim temp As IGeneric
    
    Do
        Set temp = Objects(i)
        Set Objects(i) = Objects(j)
        Set Objects(j) = temp
    Loop While InterlockedIncrement(i) < InterlockedDecrement(j)

End Sub

Public Function Median(ByVal lo As Long, ByVal hi As Long) As Single
    If (lo >= hi) Then _
        Call Me.Throw(Me, "Public Function Median").ArgOutOfRange("lo <= hi")
    
    Median = lo + (Me.RightShift(hi - lo, 1))
End Function

Public Function HashValue( _
                            ByRef ascii() As Byte, _
                            Optional ByVal Hashfunction As Hashing = Hashing.Function1) As Long
    
    Dim h As Double
    Dim i As Long
    
    Select Case Hashfunction
        Case Hashing.Function1
            
            Const Shift As Long = 2& ^ 5
            Dim h1 As Double: h1 = 352654597# 'LeftShift(5381,16))+5381
            Dim h2 As Double: h2 = h1
            With Me
                For i = LBound(ascii) To UBound(ascii) - 1
                    h1 = .X0R((h1 * Shift) + h1, ascii(i)): h2 = .X0R((h2 * Shift) + h2, ascii(i + 1))
                Next
            End With
            
            h = h1 + (h2 * 1566083941#)
        
        Case Hashing.Function2
            
             Dim Length As Long: Length = UBound(ascii) + 1
             Dim N As Long
             
             For N = (Length / 2) To 1 Step -1
                h = h + ascii(i)
                h = Me.X0R(Me.X0R(leftShift(h, 16), Me.leftShift(ascii(i + 1), 11)), h)
                h = h + Me.RightShift(h, 11)
                i = i + 2
            Next
            
            If ((Length Mod 2) = 1) Then
                h = h + ascii(i) + 1566083941
                h = Me.X0R(h, leftShift(h, 10))
                h = h + Me.RightShift(h, 1)
            End If
            
            h = Me.X0R(h, Me.leftShift(h, 3)): h = h + Me.RightShift(h, 5)
            h = Me.X0R(h, Me.leftShift(h, 4)): h = h + Me.RightShift(h, 17)
            h = Me.X0R(h, Me.leftShift(h, 25)): h = h + Me.RightShift(h, 6)
    End Select
    
    HashValue = CLng(h - (Fix(h / MAXVALUE) * MAXVALUE))

End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
End Sub

Public Function IteratorOf(ByVal c As IGenericCollection) As IGenericIterator: Set IteratorOf = c.Iterator: End Function
Public Function ClassName(ByRef Instance As IGeneric) As String: ClassName = "<" & TypeName$(Instance) & ">": End Function
Public Sub Dispose(ByVal Instance As IGeneric): Call Instance.Dispose: End Sub
Public Function Clone(ByVal Instance As IGeneric) As IGeneric: Set Clone = Instance.Clone: End Function
Public Function Generic(ByRef Instance As IGeneric) As IGeneric: Set Generic = Instance: End Function
Public Function Modulo(ByVal a As Double, ByVal m As Double) As Long: Modulo = (a - (Int(a / m) * m)): End Function
Public Function Decrement(ByRef i As Long) As Long: i = (Not -i): Decrement = i: End Function
Public Function Increment(ByRef i As Long) As Long: i = (-(Not i)): Increment = i: End Function
Public Function RightShift(ByVal Value As Long, Shift As Byte) As Double: RightShift = Value / (2& ^ Shift): End Function
Public Function leftShift(ByVal Value As Long, Shift As Byte) As Double: leftShift = Value * (2& ^ Shift): End Function
Public Function LimitDouble(ByVal d As Double) As Long: LimitDouble = CLng(d - (Fix(d / MAXVALUE) * MAXVALUE)): End Function
Public Function X0R(ByVal d1 As Double, ByVal d2 As Double) As Long: X0R = CLng(d1 - (Fix(d1 / MAXVALUE) * MAXVALUE)) Xor CLng(d2 - (Fix(d2 / MAXVALUE) * MAXVALUE)): End Function
Public Function Logn(ByVal Value, Optional ByVal Base As Byte = 2) As Long: Logn = Log(Value) / Log(Base): End Function

Public Sub ForEach(ByVal Iterator As IGenericIterator)
    
    Dim item As IGeneric
    With Iterator
        Do While .HasNext(item): RaiseEvent Delegate(item): Loop
    End With

End Sub
Public Sub Delegate(ByRef Object As IGeneric): RaiseEvent Delegate(Object): End Sub

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Range.Current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    With myclass.Range
        If .i = .Last Then _
            Set .Current = Nothing: Set Out = Nothing: Exit Function
            
        Set .Current = GNumeric(InterlockedIncrement(.i)): Set Out = .Current
        IGenericIterator_HasNext = True
      
    End With
    
End Function

Private Sub IGenericIterator_Reset()
    Set myclass.Range.Current = Nothing
    myclass.Range.i = myclass.Range.Start
End Sub
