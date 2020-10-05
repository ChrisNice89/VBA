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
Implements IPrototype

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
    Private Declare Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Size As Long, ByRef dst As LongPtr, ByVal src As Long) As Long
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

Public Enum SortOrder
    descending = -1
    random = 0
    ascending = 1
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
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get SortOrder() As SortOrder: SortOrder = DEFAULT_SORTORDER: End Property
Public Property Get PointerLength() As Byte: PointerLength = POINTERSIZE: End Property
Public Property Get MaxLong() As Long: MaxLong = MAX_INT32: End Property
Public Property Get MaxHash() As Double: MaxHash = MAXVALUE: End Property

Public Property Get Range(ByVal Start As Long, ByVal Last As Long) As IGenericIterator
    
    Dim C As Member
    
    If Start > Last Then _
        Call Me.Throw(Me, "Public Function Range").InvalidInput("Start > Last")
        
    With C.Range
        .Start = Start
        .Last = Last
        .i = .Start - 1
    End With
   
    Set Range = Skynet.CreateInstance(New Skynet, VarPtr(C), LenB(C))

End Property

Public Function Throw(ByVal Source As Object, ByVal Method As String) As Throwhelper
    Set Throw = Throwhelper.Build(Source, Method)
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

Public Function CreateInstance(ByVal Prototype As IPrototype, ByVal Source As LongPtr, ByVal Size As Long) As Object
        
    With Prototype
        If (Prototype Is .StaticInstance) Then _
            Set Prototype = .CreateDefault
        'Call CopyBytes(.Size, ByVal Clone.VirtualPointer, ByVal .VirtualPointer) 'works
        Call CopyMemory_byVar(ByVal .VirtualPointer, ByVal Source, Size): Call ZeroMemory(ByVal Source, Size)
    End With

    Set CreateInstance = Prototype

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
        
        Dim i As Long, C As ComparerEnum
        Do While (InterlockedIncrement(i) < a.Count And C = IsEqual)
            If Not a(i) Is Nothing Then _
                C = a(i).Equals(b(i))
        Loop
    End If
    
    SequenceCompare = C
    
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
                                ByRef Objects() As IGeneric, _
                                ByVal Value As IGeneric, _
                                ByVal Index As Long, _
                                ByVal Length As Long, _
                                ByVal Order As SortOrder) As Long

    Dim i As Long
    Dim C As ComparerEnum
    Dim lo As Long: lo = Index
    Dim hi As Long: hi = Index + Length - 1
    
    Do While (lo <= hi)
    
        i = lo + ((hi - lo) / 2)
      
        If (Objects(i) Is Nothing) = False Then
            C = Objects(i).CompareTo(Value) * Order
        Else
            C = IIf(Value Is Nothing, ComparerEnum.IsEqual, ComparerEnum.IsLower * Order)
        End If
        
        If (C = IsEqual) Then _
            BinarySearch = i: Exit Function
        
        If (C = IsLower) Then
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
                        ByVal Index As Long, _
                        ByVal Length As Long)

    Dim i As Long: i = Index
    Dim j As Long: j = Index + Length - 1
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
    
    Dim H As Double
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
            
            H = h1 + (h2 * 1566083941#)
        
        Case Hashing.Function2
            
             Dim Length As Long: Length = UBound(ascii) + 1
             Dim n As Long
             
             For n = (Length / 2) To 1 Step -1
                H = H + ascii(i)
                H = Me.X0R(Me.X0R(LeftShift(H, 16), Me.LeftShift(ascii(i + 1), 11)), H)
                H = H + Me.RightShift(H, 11)
                i = i + 2
            Next
            
            If ((Length Mod 2) = 1) Then
                H = H + ascii(i) + 1566083941
                H = Me.X0R(H, LeftShift(H, 10))
                H = H + Me.RightShift(H, 1)
            End If
            
            H = Me.X0R(H, Me.LeftShift(H, 3)): H = H + Me.RightShift(H, 5)
            H = Me.X0R(H, Me.LeftShift(H, 4)): H = H + Me.RightShift(H, 17)
            H = Me.X0R(H, Me.LeftShift(H, 25)): H = H + Me.RightShift(H, 6)
    End Select
    
    HashValue = CLng(H - (Fix(H / MAXVALUE) * MAXVALUE))

End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Virtual.Size = LenB(myClass)
End Sub

Public Function IteratorOf(ByVal C As IGenericCollection) As IGenericIterator: Set IteratorOf = C.Iterator: End Function
Public Function ClassName(ByRef Instance As IGeneric) As String: ClassName = "<" & TypeName$(Instance) & ">": End Function
Public Sub Dispose(ByVal Instance As IGeneric): Call Instance.Dispose: End Sub
Public Function Clone(ByVal Instance As IGeneric) As IGeneric: Set Clone = Instance.Clone: End Function
Public Function Generic(ByRef Instance As IGeneric) As IGeneric: Set Generic = Instance: End Function
Public Function Modulo(ByVal a As Double, ByVal m As Double) As Long: Modulo = (a - (Int(a / m) * m)): End Function
Public Function Decrement(ByRef i As Long) As Long: i = (Not -i): Decrement = i: End Function
Public Function Increment(ByRef i As Long) As Long: i = (-(Not i)): Increment = i: End Function
Public Function RightShift(ByVal Value As Long, Shift As Byte) As Double: RightShift = Value / (2& ^ Shift): End Function
Public Function LeftShift(ByVal Value As Long, Shift As Byte) As Double: LeftShift = Value * (2& ^ Shift): End Function
Public Function LimitDouble(ByVal d As Double) As Long: LimitDouble = CLng(d - (Fix(d / MAXVALUE) * MAXVALUE)): End Function
Public Function X0R(ByVal d1 As Double, ByVal d2 As Double) As Long: X0R = CLng(d1 - (Fix(d1 / MAXVALUE) * MAXVALUE)) Xor CLng(d2 - (Fix(d2 / MAXVALUE) * MAXVALUE)): End Function
Public Function Logn(ByVal Value, Optional ByVal Base As Byte = 2) As Long: Logn = Log(Value) / Log(Base): End Function

Public Sub ForEach(ByVal Iterator As IGenericIterator)
    
    Dim Item As IGeneric
    With Iterator
        Do While .HasNext(Item): RaiseEvent Delegate(Item): Loop
    End With

End Sub
Public Sub Delegate(ByRef Object As IGeneric): RaiseEvent Delegate(Object): End Sub

'Sorting
'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub Sort(ByRef Objects() As IGeneric, ByVal Index As Long, ByVal Length As Long, Optional ByVal Order As SortOrder = DEFAULT_SORTORDER)
    
    Const QuickSortDepthThreshold As Long = 32
    
    If Order = random Then
        Call Shuffle(Objects, Index, Length)
    ElseIf Order = ascending Then
        Call DepthLimitedQuicksort_Asc(Objects, Index, Length, QuickSortDepthThreshold)
    Else
        Call DepthLimitedQuicksort_Desc(Objects, Index, Length, QuickSortDepthThreshold)
    End If
   
End Sub

Public Sub SortKeysWithItems( _
                                ByRef Keys() As IGeneric, _
                                ByRef Items() As Object, _
                                ByVal Index As Long, _
                                ByVal Length As Long, _
                                Optional ByVal Order As SortOrder = DEFAULT_SORTORDER)
    
    Const QuickSortDepthThreshold As Long = 32
    
    If Order = random Then
        Call ShuffleWithItems(Keys, Items, Index, Length)
    ElseIf Order = ascending Then
        Call DepthLimitedQuicksortWithItems_Asc(Keys, Items, Index, Length, QuickSortDepthThreshold)
    Else
        Call DepthLimitedQuicksortWithItems_Desc(Keys, Items, Index, Length, QuickSortDepthThreshold)
    End If
   
End Sub

Public Sub SortGenericArrays( _
                                ByVal Keys As GenericArray, _
                                ByVal Items As GenericArray, _
                                ByVal Index As Long, _
                                ByVal Length As Long, _
                                Optional ByVal Order As SortOrder = DEFAULT_SORTORDER)

    Dim KeyArray() As IGeneric
    Dim ItemArray() As Object
    
    Dim d1 As Long, d2 As Long
    Dim ptrKeys As LongPtr, ptrItems As LongPtr
    
    If Keys Is Nothing Or Items Is Nothing Then _
        Call Skynet.Throw(Me, "Public Sub SortGenericArrays").NullPointer("Keys Is Nothing Or Items Is Nothing")
    
    d1 = Keys.Descriptor
    d2 = Items.Descriptor
    
    If (d1 = 0) Or (d2 = 0) Then _
        Call Skynet.Throw(Me, "Public Sub SortGenericArrays").NullPointer("Array.Descriptor = 0")
    
    ptrKeys = VarPtrArray(KeyArray)
    ptrItems = VarPtrArray(ItemArray)
    
    With Me
        Call .CopyData(ByVal ptrKeys, VarPtr(d1), POINTERSIZE): Call .CopyData(ByVal ptrItems, VarPtr(d2), POINTERSIZE)
        Call .SortKeysWithItems(KeyArray, ItemArray, Index, Length, Order)
        Call .ClearData(ByVal ptrKeys, POINTERSIZE): Call .ClearData(ByVal ptrItems, POINTERSIZE)
    End With
End Sub

Private Sub Shuffle(ByRef Element() As IGeneric, ByVal Index As Long, ByVal Length As Long)

    Dim i As Long
    Randomize VBA.Timer
    
    For i = Length To Index Step -1
        Call Swap(Element, i, (i * Rnd + 0.5))
    Next

End Sub

Private Sub ShuffleWithItems(ByRef Keys() As IGeneric, ByRef Items() As Object, ByVal Index As Long, ByVal Length As Long)

    Dim i As Long
    Randomize VBA.Timer
    
    For i = Length To Index Step -1
        Call SwapWithItem(Keys, Items, i, (i * Rnd + 0.5))
    Next

End Sub

Private Sub DepthLimitedQuicksort_Asc( _
                                         ByRef Element() As IGeneric, _
                                         ByVal Left As Long, _
                                         ByVal Right As Long, _
                                         ByVal DepthLimit As Long)

    Dim i As Long
    Dim j As Long
    Dim middle As Long

    Dim pivot As IGeneric
    Dim tmp As IGeneric

    If (Left > Right) Then _
        Exit Sub
   
    Do
        If (DepthLimit = 0) Then _
            Call HeapSort_Asc(Element, Left, Right): Exit Sub

        i = Left: j = Right
        middle = i + ((j - i) / 2)

        Call SwapIfGreater(Element, i, middle)
        Call SwapIfGreater(Element, i, j)
        Call SwapIfGreater(Element, middle, j)

        Set pivot = Element(middle)

        Do
    
            If (pivot Is Nothing) Then
                Do While Not (Element(j) Is Nothing): j = j - 1: Loop
            Else
                Do While (pivot.CompareTo(Element(i)) = IsGreater): i = i + 1: Loop
                Do While (pivot.CompareTo(Element(j)) = IsLower): j = j - 1: Loop
            End If

            If (i > j) Then _
                Exit Do

            If (i < j) Then _
                Set tmp = Element(i): _
                Set Element(i) = Element(j): _
                Set Element(j) = tmp

            i = i + 1: j = j - 1

        Loop While (i <= j)

        DepthLimit = DepthLimit - 1

        If (j - Left) <= (Right - i) Then
            If (Left < j) Then _
                Call DepthLimitedQuicksort_Asc(Element, Left, j, DepthLimit)
            Left = i
        Else
            If (i < Right) Then _
                Call DepthLimitedQuicksort_Asc(Element, i, Right, DepthLimit)
            Right = j
        End If

    Loop While (Left < Right)

End Sub
                              
Private Sub DepthLimitedQuicksort_Desc( _
                                          ByRef Element() As IGeneric, _
                                          ByVal Left As Long, _
                                          ByVal Right As Long, _
                                          ByVal DepthLimit As Long)

    Dim i As Long
    Dim j As Long
    Dim middle As Long

    Dim pivot As IGeneric
    Dim tmp As IGeneric

    If (Left > Right) Then _
        Exit Sub
    
    Do
        If (DepthLimit = 0) Then _
            Call HeapSort_Desc(Element, Left, Right): Exit Sub

        i = Left: j = Right
        middle = i + ((j - i) / 2)

        Call SwapIfGreater(Element, i, middle)
        Call SwapIfGreater(Element, i, j)
        Call SwapIfGreater(Element, middle, j)

        Set pivot = Element(middle)

        Do
        
            If (pivot Is Nothing) Then
                Do While Not (Element(i) Is Nothing): i = i + 1: Loop
            Else
                Do While (pivot.CompareTo(Element(i)) = IsLower): i = i + 1: Loop
                Do While (pivot.CompareTo(Element(j)) = IsGreater): j = j - 1: Loop
            End If

            If (i > j) Then _
                Exit Do
        
            If (i < j) Then _
                Set tmp = Element(i): Set Element(i) = Element(j): Set Element(j) = tmp
        
            i = i + 1: j = j - 1
            
        Loop While (i <= j)

        DepthLimit = DepthLimit - 1

        If (j - Left) <= (Right - i) Then
            If (Left < j) Then _
                Call DepthLimitedQuicksort_Desc(Element, Left, j, DepthLimit)
            Left = i
        Else
            If (i < Right) Then _
                Call DepthLimitedQuicksort_Desc(Element, i, Right, DepthLimit)
            Right = j
        End If

    Loop While (Left < Right)

End Sub

Private Sub HeapSort_Asc( _
                            ByRef Keys() As IGeneric, _
                            ByVal lo As Long, _
                            ByVal hi As Long)

    Dim n As Long: n = hi - lo + 1
    Dim i As Long
 
    For i = n / 2 To 1 Step -1
        Call MaxHeap(Keys, i, n, lo)
    Next
    For i = n To 2 Step -1
        Call Swap(Keys, lo, lo + i - 1)
        Call MaxHeap(Keys, 1, i - 1, lo)
    Next
End Sub

Private Sub MaxHeap(ByRef Keys() As IGeneric, ByVal i As Long, ByVal n As Long, ByVal lo As Long)

    Dim d As IGeneric: Set d = Keys(lo + i - 1)
    Dim Child As Long

    Do While (i <= n / 2)
        Child = 2 * i
        If (Child < n) Then
            If (Keys(lo + Child - 1) Is Nothing) Then
                Child = Child + 1
            ElseIf (Keys(lo + Child - 1).CompareTo(Keys(lo + Child)) = IsLower) Then
                Child = Child + 1
            End If
        End If

        If (Keys(lo + Child - 1) Is Nothing) Then _
            Exit Do

        If (Keys(lo + Child - 1).CompareTo(d) = IsLower) Then _
            Exit Do

        Set Keys(lo + i - 1) = Keys(lo + Child - 1)
        i = Child
    Loop

    Set Keys(lo + i - 1) = d

End Sub

Private Sub HeapSort_Desc( _
                             ByRef Keys() As IGeneric, _
                             ByVal lo As Long, _
                             ByVal hi As Long)

    Dim n As Long: n = hi - lo + 1
    Dim i As Long
 
    For i = n / 2 To 1 Step -1
        Call MinHeap(Keys, i, n, lo)
    Next
    For i = n To 2 Step -1
        Call Swap(Keys, lo, lo + i - 1)
        Call MinHeap(Keys, 1, i - 1, lo)
    Next
End Sub

Private Sub MinHeap(ByRef Keys() As IGeneric, ByVal i As Long, ByVal n As Long, ByVal lo As Long)

    Dim d As IGeneric: Set d = Keys(lo + i - 1)
    Dim Child As Long

    Do While (i <= n / 2)
        Child = 2 * i
        If (Child < n) Then
            If Not (Keys(lo + Child - 1) Is Nothing) Then
                If (Keys(lo + Child - 1).CompareTo(Keys(lo + Child)) = IsGreater) Then _
                    Child = Child + 1
    
            End If
        End If

        If Not (Keys(lo + Child - 1) Is Nothing) Then
            If (Keys(lo + Child - 1).CompareTo(d) = IsGreater) Then _
                Exit Do
        
        End If
        
        Set Keys(lo + i - 1) = Keys(lo + Child - 1)
        i = Child
    Loop

    Set Keys(lo + i - 1) = d

End Sub

Private Sub DepthLimitedQuicksortWithItems_Asc( _
                                                  ByRef Keys() As IGeneric, _
                                                  ByRef Items() As Object, _
                                                  ByVal Left As Long, _
                                                  ByVal Right As Long, _
                                                  ByVal DepthLimit As Long)

    Dim i As Long
    Dim j As Long
    Dim middle As Long

    Dim pivot As IGeneric
    Dim tmp As IGeneric

    If (Left > Right) Then _
        Exit Sub
   
    Do
        If (DepthLimit = 0) Then _
            Call HeapSort_Asc(Keys, Left, Right): Exit Sub

        i = Left: j = Right
        middle = i + ((j - i) / 2)

        Call SwapIfGreaterWithItem(Keys, Items, i, middle)
        Call SwapIfGreaterWithItem(Keys, Items, i, j)
        Call SwapIfGreaterWithItem(Keys, Items, middle, j)

        Set pivot = Keys(middle)

        Do
    
            If (pivot Is Nothing) Then
                Do While Not (Keys(j) Is Nothing): j = j - 1: Loop
            Else
                Do While (pivot.CompareTo(Keys(i)) = IsGreater): i = i + 1: Loop
                Do While (pivot.CompareTo(Keys(j)) = IsLower): j = j - 1: Loop
            End If

            If (i > j) Then _
                Exit Do

            If (i < j) Then _
                Set tmp = Keys(i): Set Keys(i) = Keys(j): Set Keys(j) = tmp: _
                Set tmp = Items(i): Set Items(i) = Items(j): Set Items(j) = tmp
                
            i = i + 1: j = j - 1

        Loop While (i <= j)

        DepthLimit = DepthLimit - 1

        If (j - Left) <= (Right - i) Then
            If (Left < j) Then _
                Call DepthLimitedQuicksortWithItems_Asc(Keys, Items, Left, j, DepthLimit)
            Left = i
        Else
            If (i < Right) Then _
                Call DepthLimitedQuicksortWithItems_Asc(Keys, Items, i, Right, DepthLimit)
            Right = j
        End If

    Loop While (Left < Right)

End Sub
                              
Private Sub DepthLimitedQuicksortWithItems_Desc( _
                                                   ByRef Keys() As IGeneric, _
                                                   ByRef Items() As Object, _
                                                   ByVal Left As Long, _
                                                   ByVal Right As Long, _
                                                   ByVal DepthLimit As Long)

    Dim i As Long
    Dim j As Long
    Dim middle As Long

    Dim pivot As IGeneric
    Dim tmp As IGeneric

    If (Left > Right) Then _
        Exit Sub
    
    Do
        If (DepthLimit = 0) Then _
            Call HeapSort_Desc(Keys, Left, Right): Exit Sub

        i = Left: j = Right
        middle = i + ((j - i) / 2)

        Call SwapIfGreaterWithItem(Keys, Items, i, middle)
        Call SwapIfGreaterWithItem(Keys, Items, i, j)
        Call SwapIfGreaterWithItem(Keys, Items, middle, j)

        Set pivot = Keys(middle)

        Do
        
            If (pivot Is Nothing) Then
                Do While Not (Keys(i) Is Nothing): i = i + 1: Loop
            Else
                Do While (pivot.CompareTo(Keys(i)) = IsLower): i = i + 1: Loop
                Do While (pivot.CompareTo(Keys(j)) = IsGreater): j = j - 1: Loop
            End If

            If (i > j) Then _
                Exit Do
        
            If (i < j) Then _
                Set tmp = Keys(i): Set Keys(i) = Keys(j): Set Keys(j) = tmp: _
                Set tmp = Items(i): Set Items(i) = Items(j): Set Items(j) = tmp
        
            i = i + 1: j = j - 1
            
        Loop While (i <= j)

        DepthLimit = DepthLimit - 1

        If (j - Left) <= (Right - i) Then
            If (Left < j) Then _
                Call DepthLimitedQuicksortWithItems_Desc(Keys, Items, Left, j, DepthLimit)
            Left = i
        Else
            If (i < Right) Then _
                Call DepthLimitedQuicksortWithItems_Desc(Keys, Items, i, Right, DepthLimit)
            Right = j
        End If

    Loop While (Left < Right)

End Sub

Private Sub HeapSortWithItems_Asc( _
                                     ByRef Keys() As IGeneric, _
                                     ByRef Items() As Object, _
                                     ByVal lo As Long, _
                                     ByVal hi As Long)

    Dim n As Long: n = hi - lo + 1
    Dim i As Long
 
    For i = n / 2 To 1 Step -1
        Call MaxHeapWithItems(Keys, Items, i, n, lo)
    Next
    For i = n To 2 Step -1
        Call SwapWithItem(Keys, Items, lo, lo + i - 1)
        Call MaxHeapWithItems(Keys, Items, 1, i - 1, lo)
    Next
End Sub

Private Sub MaxHeapWithItems( _
                                ByRef Keys() As IGeneric, _
                                ByRef Items() As Object, _
                                ByVal i As Long, _
                                ByVal n As Long, _
                                ByVal lo As Long)

    Dim d As IGeneric: Set d = Keys(lo + i - 1)
    Dim dItem As Object: Set dItem = Items(lo + i - 1)
    Dim Child As Long

    Do While (i <= n / 2)
        Child = 2 * i
        If (Child < n) Then
            If (Keys(lo + Child - 1) Is Nothing) Then
                Child = Child + 1
            ElseIf (Keys(lo + Child - 1).CompareTo(Keys(lo + Child)) = IsLower) Then
                Child = Child + 1
            End If
        End If

        If (Keys(lo + Child - 1) Is Nothing) Then _
            Exit Do

        If (Keys(lo + Child - 1).CompareTo(d) = IsLower) Then _
            Exit Do

        Set Keys(lo + i - 1) = Keys(lo + Child - 1)
        Set Items(lo + i - 1) = Items(lo + Child - 1)
        i = Child
    Loop

    Set Keys(lo + i - 1) = d
    Set Items(lo + i - 1) = dItem

End Sub

Private Sub HeapSortWithItems_Desc( _
                                        ByRef Keys() As IGeneric, _
                                        ByRef Items() As Object, _
                                        ByVal lo As Long, _
                                        ByVal hi As Long)

    Dim n As Long: n = hi - lo + 1
    Dim i As Long
 
    For i = n / 2 To 1 Step -1
        Call MinHeapWithItems(Keys, Items, i, n, lo)
    Next
    For i = n To 2 Step -1
        Call SwapWithItem(Keys, Items, lo, lo + i - 1)
        Call MinHeapWithItems(Keys, Items, 1, i - 1, lo)
    Next
End Sub

Private Sub MinHeapWithItems( _
                                ByRef Keys() As IGeneric, _
                                ByRef Items() As Object, _
                                ByVal i As Long, _
                                ByVal n As Long, _
                                ByVal lo As Long)

    Dim d As IGeneric: Set d = Keys(lo + i - 1)
    Dim dItem As Object: Set dItem = Items(lo + i - 1)
    Dim Child As Long

    Do While (i <= n / 2)
        Child = 2 * i
        If (Child < n) Then
            If Not (Keys(lo + Child - 1) Is Nothing) Then
                If (Keys(lo + Child - 1).CompareTo(Keys(lo + Child)) = IsGreater) Then _
                    Child = Child + 1
    
            End If
        End If

        If Not (Keys(lo + Child - 1) Is Nothing) Then
            If (Keys(lo + Child - 1).CompareTo(d) = IsGreater) Then _
                Exit Do
        
        End If
        
        Set Keys(lo + i - 1) = Keys(lo + Child - 1)
        Set Items(lo + i - 1) = Items(lo + Child - 1)
        
        i = Child
    Loop

    Set Keys(lo + i - 1) = d
    Set Items(lo + i - 1) = dItem
    
End Sub

Private Sub SwapIfGreater(ByRef Keys() As IGeneric, ByVal a As Long, ByVal b As Long)
    
    Dim T As IGeneric
   
    If Not (a = b) Then
        If Not (Keys(a) Is Nothing) Then
            If (Keys(a).CompareTo(Keys(b)) = IsGreater) Then
                Set T = Keys(a)
                Set Keys(a) = Keys(b)
                Set Keys(b) = T
            End If
        End If
    End If
    
End Sub

Private Sub SwapIfGreaterWithItem(ByRef Keys() As IGeneric, ByRef Items() As Object, ByVal a As Long, ByVal b As Long)
    
    Dim T As IGeneric
   
    If Not (a = b) Then
        If Not (Keys(a) Is Nothing) Then
            If (Keys(a).CompareTo(Keys(b)) = IsGreater) Then
                Set T = Keys(a): Set Keys(a) = Keys(b): Set Keys(b) = T
                Set T = Items(a): Set Items(a) = Items(b): Set Items(b) = T
            End If
        End If
    End If
    
End Sub

Private Sub Swap(ByRef a() As IGeneric, ByVal i As Long, ByVal j As Long)

    Dim T As IGeneric
    
    Set T = a(i): Set a(i) = a(j): Set a(j) = T

End Sub

Private Sub SwapWithItem(ByRef a() As IGeneric, ByRef b() As Object, ByVal i As Long, ByVal j As Long)
    
    Dim T As IGeneric
    
    Set T = a(i): Set a(i) = a(j): Set a(j) = T
    Set T = b(i): Set b(i) = b(j): Set b(j) = T
    
End Sub

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myClass.Range.Current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    With myClass.Range
        If .i = .Last Then _
            Set .Current = Nothing: Set Out = Nothing: Exit Function
            
        Set .Current = GNumeric(InterlockedIncrement(.i)): Set Out = .Current
        IGenericIterator_HasNext = True
      
    End With
    
End Function

Private Sub IGenericIterator_Reset()
    Set myClass.Range.Current = Nothing
    myClass.Range.i = myClass.Range.Start
End Sub

'IPrototype
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New Skynet: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = Skynet: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.Size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property


