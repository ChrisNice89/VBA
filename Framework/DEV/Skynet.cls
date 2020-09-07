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

Private Const DEFAULT_SORTORDER As Integer = 1 'SortOrder.Ascending
Private Const MAXVALUE As Double = 2 ^ 31
Private Const MAX_INT32 As Long = &H7FFFFFFF
Private Const NullPointer As Long = 0

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
    Random = 0
    Ascending = 1
End Enum

Public Enum SortMethod
    IntrospectiveSort = -1
    DepthLimitedQuicksort = 1
End Enum

Public Enum SystemState
    Ready
    BuildNew
End Enum

Private Type Member
    x As Long
End Type
Private myClass As Member
'@MemberAttribute VB_VarHelpID, -1

Private WithEvents WordSequence As GenericEnumerable
Attribute WordSequence.VB_VarHelpID = -1
Private WithEvents IntegerSequence As GenericEnumerable
Attribute IntegerSequence.VB_VarHelpID = -1
Private WithEvents ContainsEnumerable As GenericEnumerable
Attribute ContainsEnumerable.VB_VarHelpID = -1

Public Property Get SortOrder() As SortOrder: SortOrder = DEFAULT_SORTORDER: End Property
Public Property Get PointerLength() As Byte: PointerLength = POINTERSIZE: End Property
Public Property Get MaxLong() As Long: MaxLong = MAX_INT32: End Property

Public Function Throw(ByVal Source As Object, ByVal Method As String) As Throwhelper
    Set Throw = Throwhelper.Build(Source, Method)
End Function

Public Sub ClearData(ByVal Dst As LongPtr, ByVal Bytes As Long)
    Call FillMemory(ByVal Dst, Bytes, 0) 'Call ZeroMemory(ByVal dst, Bytes)
End Sub

Public Sub CopyData(ByVal Dst As LongPtr, ByVal Src As LongPtr, ByVal Bytes As Long, Optional ByVal ClearSource As Boolean = False)
    
    Call CopyMemory_byVar(ByVal Dst, ByVal Src, Bytes) 'Works
    'Call CopyBytes(Bytes, ByVal Dst, ByVal Src) 'works
    If ClearSource Then _
        Call ZeroMemory(ByVal Src, Bytes) 'FillMemory(ByVal Src, Bytes, 0)
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
    
        Call .CopyData(Dst:=ByVal VarPtr(vType), Src:=ByVal VarPtr(SomeArray), Bytes:=2&)
        Call .CopyData(Dst:=ByVal VarPtr(PointerOf), Src:=ByVal VarPtr(SomeArray) + 8&, Bytes:=POINTERSIZE)
        
        'if SomeArray is defined x() as ...
        If (vType And VT_BYREF) = VT_BYREF Then _
            Call .CopyData(Dst:=ByVal VarPtr(PointerOf), Src:=ByVal PointerOf, Bytes:=POINTERSIZE)
    End With
    
End Function

'SequenceCompare
Public Function SequenceCompare(ByVal a As IGenericReadOnlyList, ByVal b As IGenericReadOnlyList) As ComparerEnum
    
    If (a Is b) Then SequenceCompare = IsEqual: Exit Function
    
    If Not (a Is Nothing) And (b Is Nothing) Then
        
        If Not (Me.Generic(a).IsRelatedTo(b)) Then _
            Me.Throw(Me, "CompareLists").InvalidInput ("Type(a) != Type(b)")
    
        If Not (a.Elements.Count = b.Elements.Count) Then _
            Me.Throw(Me, "CompareLists").InvalidInput ("a.Elements.Count != b.Elements.Count")
        
        Dim i As Long, C As ComparerEnum
        Do While (InterlockedIncrement(i) < a.Elements.Count And C = IsEqual)
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
      
        If (Objects(i) Is Nothing) Then
            C = IIf(Value Is Nothing, ComparerEnum.IsEqual, ComparerEnum.IsLower * Order)
        Else
            C = Objects(i).CompareTo(Value) * Order
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
                            ByVal Count As Long)
    Dim i As Long
    Dim EndIndex As Long

    EndIndex = StartIndex + Count - LBound(Objects)
    
    If (Value Is Nothing) Then
        For i = StartIndex To EndIndex
            If (Objects(i) Is Nothing) Then _
                IndexOf = i: Exit Function
        Next
    Else
        For i = StartIndex To EndIndex
            If Value.Equals(Objects(i)) Then _
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

'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub Sort(ByRef Objects() As IGeneric, ByVal Index As Long, ByVal Length As Long, Optional ByVal Order As SortOrder = DEFAULT_SORTORDER)
    
    Const QuickSortDepthThreshold As Long = 32
    
    If Order = Random Then
        Call Shuffle(Objects)
    ElseIf Order = Ascending Then
        Call DepthLimitedQuicksort_Asc(Objects, Index, Length, QuickSortDepthThreshold)
    Else
        Call DepthLimitedQuicksort_Desc(Objects, Index, Length, QuickSortDepthThreshold)
    End If
   
End Sub

Private Sub Shuffle(ByRef Element() As IGeneric)

    Dim i As Long
    Randomize VBA.Timer
    
    For i = UBound(Element) To LBound(Element) Step -1
        Call Swap(Element, i, (i * Rnd + 0.5))
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
                Set tmp = Element(i): _
                Set Element(i) = Element(j): _
                Set Element(j) = tmp

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
Private Sub SwapIfGreater(ByRef Keys() As IGeneric, ByVal a As Long, ByVal b As Long)
    
    Dim t As IGeneric
   
    If Not (a = b) Then
        If Not (Keys(a) Is Nothing) Then
            If (Keys(a).CompareTo(Keys(b)) = IsGreater) Then
                Set t = Keys(a)
                Set Keys(a) = Keys(b)
                Set Keys(b) = t
            End If
        End If
    End If
    
End Sub

Private Sub Swap(ByRef a() As IGeneric, ByVal i As Long, ByVal j As Long)
    Dim t As IGeneric
    
    Set t = a(i)
    Set a(i) = a(j)
    Set a(j) = t

End Sub

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
'    If Not (Me Is Skynet) Then _
        Call Me.Throw(Me, "Private Sub Class_Initialize()").Instantiation("System class must not be instantiated")
   
End Sub

Public Function EnumeratorOf(ByVal C As IGenericCollection) As IGenericEnumerator: Set EnumeratorOf = C.NewEnumerator: End Function
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

Public Sub TestRange()

    Dim Range As IGenericEnumerator
    Set Range = GenericEnumerable.Range(-5, 5)
    
    With GenericEnumerable.Range(-5, 5)
        Do While .MoveNext
            Debug.Print .Current
        Loop
    End With
    
    Set IntegerSequence = GenericEnumerable.Build(Range)
    Debug.Print "The number of even integers is: " & IntegerSequence.Aggregate(GNumeric)
    
End Sub
    
Public Sub TestInteger2(ByVal n As Long)
    
    Dim i As Long
    
    Dim ints As GenericArray
    Set ints = GenericArray.Build(n)
    
    For i = ints.LowerBound To n
        Set ints(i) = GNumeric(i)
    Next
    
    Set IntegerSequence = GenericEnumerable.Build(ints)
    
    Debug.Print "The number of even integers is: " & IntegerSequence.Aggregate(GNumeric, IgnoreNull:=True)
    
End Sub

Public Sub TestContains()
    
    Dim i As Long
    
    Dim fruits As GenericArray
    Set fruits = GenericArray.BuildWithStrings("apple", "banana", "mango", "orange", "passionfruit", "grape")
    
    Debug.Print "Fruits: " & GString.Join(fruits, ";").Value
    
    Dim grape As GString
    Set grape = GString("grape")
    
    Dim kiwi As GString
    Set kiwi = GString("kiwi")

    Debug.Print "Has grape: " & GenericEnumerable(fruits).Contains(grape)
    Debug.Print "Has kiwi: " & GenericEnumerable(fruits).Contains(kiwi)
    
    Dim list As GenericList
    Set list = GenericList.Build(Capacity:=fruits.Length - 1)
    Call list.AddAll(fruits)
    
    Debug.Print "fruit list contains grape: " & Cast.ToCollection(list).Contains(grape)
    
End Sub

Public Sub TestString()
    
    Dim sentence As GString
    Set sentence = GString("the quick brown fox jumps over the lazy dog")
    
    Debug.Print "Before: " & sentence.Value
    
    Set WordSequence = GenericEnumerable(sentence.Split(" ").Enumerator(1, sentence.CountChar(" ") + 1))
    
    Debug.Print "After: " & WordSequence.Aggregate(GString, IgnoreNull:=True)
    
End Sub

Public Sub TestInteger()
    
    Dim ints As GenericArray
    Set ints = GenericArray.BuildWithIntegers(4, 8, 8, 3, 9, 0, 7, 8, 2)
    
    Set IntegerSequence = GenericEnumerable(ints.Enumerator(1, 9))
    Debug.Print "Integers: " & GString.Join(ints, ";").Value
    
    Debug.Print "The number of even integers is: " & IntegerSequence.Aggregate(GNumeric, IgnoreNull:=True)
    
End Sub

Private Sub IntegerSequence_Aggregate(Result As IGeneric, ByVal Current As IGeneric, ByVal NullsIgnored As Boolean)
    If (Cast.ToNumeric(Current).IsEven) Then Set Result = Cast.ToNumeric(Result).Add(1)
End Sub

Private Sub WordSequence_Aggregate(Result As IGeneric, ByVal Current As IGeneric, ByVal NullsIgnored As Boolean)
    Set Result = Cast.ToString(Current).Concat(Result, " ")
End Sub
