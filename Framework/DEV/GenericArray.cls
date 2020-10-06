VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("<T>Enumerable")


Option Explicit
'@PredeclaredId

Implements IPrototype
Implements IGeneric
Implements IGenericReadOnlyList
Implements IGenericCollection
Implements IGenericIterator

'Pointer
Private Const VT_BYREF As Long = &H4000&
Private Const S_OK As Long = &H0&
Private Const E_ABORT As Long = &H80004004
Private Const E_INVALIDARG As Long = &H80070057
Private Const E_POINTER As Long = &H80004003
Private Const DISP_E_BADINDEX As Long = &H8002000B
Private Const E_OUTOFMEMORY As Long = &H8007000E
Private Const DISP_E_ARRAYISLOCKED As Long = &H8002000D
Private Const E_UNEXPECTED As Long = &H8000FFFF

Private Const MIN_INDEX As Long = 1 'Option Base { 0 | 1 }
Private Const MAX_ARRAY_LENGHT As Long = &H7FEFFFFF
Private Const ELEMENT_SIZE As Long = 4&
Private Const ELEMENT_TYPE As Integer = 9 'vbvartype.vbObject

#If Win64 Then
    Private Declare PtrSafe Function SafeArrayCreateEx Lib "oleaut32.dll" (ByVal vt As Integer, ByVal cDims As Long, ByRef rgsabound As SAFEARRAYBOUND, ByRef pvExtra As Any) As LongPtr
    Private Declare PtrSafe Function SafeArrayDestroyDescriptor Lib "oleaut32.dll" (ByRef psa As Any) As LongPtr
    Private Declare PtrSafe Function SafeArrayDestroy Lib "oleaut32.dll" (ByRef psa As Any) As LongPtr
    Private Declare PtrSafe Function SafeArrayDestroyData Lib "oleaut32.dll" (ByRef psa As Any) As LongPtr
    Private Declare PtrSafe Function SafeArrayPutElement Lib "oleaut32.dll" (ByRef psa As Any, ByRef rgIndices As Long, ByRef pv As Any) As LongPtr
    Private Declare PtrSafe Function SafeArrayGetElement Lib "oleaut32.dll" (ByVal psa As Any, ByRef rgIndices As Long, ByRef pv As Any) As LongPtr
    Private Declare PtrSafe Function SafeArrayCopy Lib "oleaut32.dll" (ByRef psa As Any, ByRef ppsaOut As Any) As LongPtr
    Private Declare PtrSafe Function SafeArrayCopyData Lib "oleaut32.dll" (ByRef psaSource As Any, ByRef psaTarget As Any) As LongPtr
    Private Declare PtrSafe Function SafeArrayAccessData Lib "oleaut32.dll" (ByRef psa As Any, ByRef ppvData As Any) As LongPtr
    Private Declare PtrSafe Function SafeArrayUnaccessData Lib "oleaut32.dll" (ByRef psa As Any) As LongPtr
    Private Declare PtrSafe Function SafeArrayPtrOfIndex Lib "oleaut32.dll" (ByRef psa As Any, ByRef rgIndices As Long, ByRef ppvData As Any) As LongPtr
    Private Declare PtrSafe Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef psa As Any) As LongPtr
    Private Declare PtrSafe Function SafeArrayGetUBound Lib "oleaut32.dll" (ByRef psa As Any, ByVal nDim As Long, ByRef plUbound As Long) As LongPtr
    Private Declare PtrSafe Function SafeArrayRedim Lib "oleaut32.dll" (ByRef psa As Any, ByRef psaboundNew As SAFEARRAYBOUND) As Long
    Private Declare PtrSafe Function SafeArrayUnlock Lib "oleaut32.dll" (ByRef psa As Any) As Long
    Private Declare PtrSafe Function SafeArrayLock Lib "oleaut32.dll" (ByRef psa As Any) As Long
    Private Declare PtrSafe Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Var() As Any) As LongPtr
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
    Private Declare Function SafeArrayCreateEx Lib "oleaut32.dll" (ByVal vt As Integer, ByVal cDims As Long, ByRef rgsabound As SAFEARRAYBOUND, ByRef pvExtra As Any) As Long
    Private Declare Function SafeArrayDestroyDescriptor Lib "oleaut32.dll" (ByRef psa As Any) As Long
    Private Declare Function SafeArrayDestroy Lib "oleaut32.dll" (ByRef psa As Any) As Long
    Private Declare Function SafeArrayDestroyData Lib "oleaut32.dll" (ByRef psa As Any) As Long
    Private Declare Function SafeArrayPutElement Lib "oleaut32.dll" (ByRef psa As Any, ByRef rgIndices As Long, ByRef pv As Any) As Long
    Private Declare Function SafeArrayGetElement Lib "oleaut32.dll" (ByVal psa As Any, ByRef rgIndices As Long, ByRef pv As Any) As Long
    Private Declare Function SafeArrayCopy Lib "oleaut32.dll" (ByRef psa As Any, ByRef ppsaOut As Any) As Long
    Private Declare Function SafeArrayCopyData Lib "oleaut32.dll" (ByRef psaSource As Any, ByRef psaTarget As Any) As Long
    Private Declare Function SafeArrayAccessData Lib "oleaut32.dll" (ByRef psa As Any, ByRef ppvData As Any) As Long
    Private Declare Function SafeArrayUnaccessData Lib "oleaut32.dll" (ByRef psa As Any) As Long
    Private Declare Function SafeArrayPtrOfIndex Lib "oleaut32.dll" (ByRef psa As Any, ByRef rgIndices As Long, ByRef ppvData As Any) As Long
    Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef psa As Any) As Integer
    Private Declare Function SafeArrayGetUBound Lib "oleaut32.dll" (ByRef psa As Any, ByVal nDim As Long, ByRef plUbound As Long) As Long
    Private Declare Function SafeArrayRedim Lib "oleaut32.dll" (ByRef psa As Any, ByRef psaboundNew As SAFEARRAYBOUND) As Long
    Private Declare Function SafeArrayUnlock Lib "oleaut32.dll" (ByRef psa As Any) As Long
    Private Declare Function SafeArrayLock Lib "oleaut32.dll" (ByRef psa As Any) As Long
    Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Var() As Any) As Long
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

Private Type SAFEARRAYBOUND
    cElements    As Long
    lLbound      As Long
End Type
'
Private Type SAFEARRAYSTRUCT
    cDims        As Integer
    fFeatures    As Integer
    cbElements   As Long
    cLocks       As Long
    pvData       As Long
    Bounds() As SAFEARRAYBOUND
End Type

Private Type ObjectTransfer
    #If Win64 Then
        bit(1 To 8) As Byte
    #Else
        bit(1 To 4) As Byte
    #End If
End Type

Private Type Iterator
    StartIndex As Long
    EndIndex As Long
    Count As Long
    Complete As Boolean
    Rank As Long
    Indices() As Long
    Index As Long
    ptrIndices As LongPtr
    Descriptor As LongPtr
    Parent As GenericArray
    Current As IGeneric
    HasLoaded As Boolean
End Type

Private Type Member
   Struct As SAFEARRAYSTRUCT
   Descriptor As LongPtr
   Elements As Long
   Indices() As Long
   Internal() As IGeneric
   IsReadOnly As Boolean
   Iterator As Iterator
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get MaxLength() As Long: MaxLength = MAX_ARRAY_LENGHT: End Property
Public Property Get LowerBound() As Long: LowerBound = MIN_INDEX: End Property
Public Property Get IsEmpty() As Boolean: IsEmpty = (myClass.Descriptor = 0): End Property
Public Property Get Descriptor() As LongPtr: Descriptor = myClass.Descriptor: End Property
Public Property Get LengthOf(ByVal Dimension As Long) As Long: Call Catch(SafeArrayGetUBound(ByVal myClass.Descriptor, Dimension, LengthOf)): End Property
Public Property Get LengthOf_2(ByVal Dimension As Long) As Long: LengthOf_2 = myClass.Struct.Bounds(1 + myClass.Struct.cDims - Dimension).cElements: End Property
Public Property Get Length() As Long: Length = myClass.Elements: End Property
Public Property Get TotalSize() As LongPtr: TotalSize = Me.SizeOf(myClass.Elements): End Property
Public Property Get Rank() As Long: Rank = SafeArrayGetDim(ByVal myClass.Descriptor): End Property
Public Property Set ElementAt(ByVal Index As Long, ByVal Value As IGeneric): Set myClass.Internal(Index) = Value: End Property
Public Property Get ElementAt(ByVal Index As Long) As IGeneric: Set ElementAt = myClass.Internal(Index): End Property
Attribute ElementAt.VB_UserMemId = 0
' Todo
Public Function Build(ParamArray Indices() As Variant) As GenericArray
    
    Dim SAB() As SAFEARRAYBOUND
    Dim i As Long, UB As Long, LB As Long
    
    If (Me Is GenericArray) = False Then _
        Call Skynet.Throw(Me, "Public Function Build").BuildExeption
    
    If IsMissing(Indices) Then _
        Call Skynet.Throw(Me, "Public Function Build").NullPointer("ParamArray Indices are missing")
    
    LB = LBound(Indices): UB = UBound(Indices)
    
    ReDim SAB(LB To UB)
    For i = LB To UB
        If (Indices(i) < 1) Then _
            Call Skynet.Throw(Me, "Build").ArgOutOfRange("A zero(one)-based array indexed array has its first item indexed as 0(1) and need to have at least one element")
        
        With SAB(i)
            .cElements = Indices(i)
            .lLbound = MIN_INDEX
        End With
    Next i
    
    'Call Catch(SafeArrayDestroyDescriptor(ByVal Skynet.PointerOf(myClass.Internal))) '!Memoryleak?!

    With myClass
        .Descriptor = SafeArrayCreateEx(ELEMENT_TYPE, UB + 1, SAB(LB), ByVal 0&)
        
        If (.Descriptor = 0) Then _
            Call Skynet.Throw(Me, "Public Function Build").InvalidOperation("Descriptor = 0")
        
        Call LoadArrayStructure(.Struct, .Descriptor)
        ReDim .Indices(.Struct.cDims - 1)
    End With
    
    With Skynet
        Call .CopyData(ByVal VarPtrArray(myClass.Internal), VarPtr(myClass.Descriptor), .PointerLength)
        Set Build = .CreateInstance(New GenericArray, Virtual.Pointer, Virtual.Size)
    End With
    
End Function

Private Sub LoadArrayStructure(ByRef Structure As SAFEARRAYSTRUCT, ByVal SafeArrayPointer As LongPtr)
    
    Dim i As Long
    
    If (SafeArrayPointer = 0) Then _
        Skynet.Throw(Me, "Private Sub LoadArrayStructure").NullPointer ("SafeArrayPointer")

    Call Skynet.CopyData(VarPtr(Structure), SafeArrayPointer, 16&)
    
    With Structure
        If (.cDims <= 0) Then _
            Call Skynet.Throw(Me, "Private Sub LoadArrayStructure").Overflow(".cDims <= 0")
        
        ReDim .Bounds(1 To .cDims)
        Call Skynet.CopyData(VarPtr(.Bounds(1)), ByVal (SafeArrayPointer + 16&), .cDims * 8&)
        
        myClass.Elements = 1
        For i = LBound(.Bounds) To UBound(.Bounds)
            myClass.Elements = myClass.Elements * .Bounds(i).cElements
        Next
        
    End With

End Sub
'Obsolet just 4 legacy
Private Function GetArrayStructure(ByVal SafeArrayPointer As LongPtr) As SAFEARRAYSTRUCT
    
    If (SafeArrayPointer = 0) Then _
        Skynet.Throw(Me, "Private Function GetArrayStructure").NullPointer ("SafeArrayPointer")

    Call Skynet.CopyData(VarPtr(GetArrayStructure), SafeArrayPointer, 16&)
    
    With GetArrayStructure
        If (.cDims <= 0) Then _
            Call Skynet.Throw(Me, "Private Function GetArrayStructure").Overflow(".cDims <= 0")
        
        ReDim .Bounds(1 To .cDims)
        Call Skynet.CopyData(VarPtr(.Bounds(1)), ByVal (SafeArrayPointer + 16&), .cDims * 8&)
    End With

End Function

Public Function BuildFrom(ByRef Matrix() As IGeneric) As GenericArray
        
    Dim i As Long
    
    If (Me Is GenericArray) = False Then _
        Call Skynet.Throw(Me, "Public Function BuildFrom").BuildExeption
    
    If Not VarType(Matrix) = VbVarType.vbArray + ELEMENT_TYPE Then _
        Call Skynet.Throw(Me, "Public Function BuildWith").InvalidInput("Datatype {" & TypeName$(Matrix) & "} not allowed")
        
    If LBound(Matrix) < MIN_INDEX Then _
        Call Skynet.Throw(Me, "Public Function BuildFrom").InvalidOperation("LBound(Matrix) < MIN_INDEX")
 
    With Skynet
        myClass.Descriptor = .PointerOf(Matrix)

        If (myClass.Descriptor = 0) Then _
            Call .Throw(Me, "Public Function BuildFrom").InvalidOperation("Descriptor = 0")
        
        Call LoadArrayStructure(myClass.Struct, myClass.Descriptor)
        
        If Not (myClass.Struct.cbElements = ELEMENT_SIZE) Then _
            Call .Throw(Me, "Public Function BuildFrom").InvalidOperation(".Struct.cbElements != ELEMENT_SIZE")
        
        Call .ClearData(ByVal VarPtrArray(Matrix), .PointerLength) '!Memoryleak?!
        Call .CopyData(ByVal VarPtrArray(myClass.Internal), VarPtr(myClass.Descriptor), .PointerLength)
        
        Set BuildFrom = .CreateInstance(New GenericArray, Virtual.Pointer, Virtual.Size)
    
    End With

End Function

Public Function BuildWith(ParamArray GenericValues() As Variant) As GenericArray
    
    Dim i As Long
    
    If (Me Is GenericArray) = False Then _
        Call Skynet.Throw(Me, "Public Function BuildWith").BuildExeption
    
    If IsMissing(GenericValues) Then _
        Call Skynet.Throw(Me, "Public Function BuildWith(").NullPointer("ParamArray GenericValues are missing")
    
    ReDim Values(MIN_INDEX To UBound(GenericValues) + MIN_INDEX) As IGeneric
    
    For i = LBound(GenericValues) To UBound(GenericValues)
        
        If Not VarType(GenericValues(i)) = ELEMENT_TYPE Then _
            Call Skynet.Throw(Me, "Public Function BuildWith").InvalidInput("Datatype {" & TypeName$(GenericValues(i)) & "} not allowed")
        
        If (GenericValues(i) Is Nothing) Then _
            Call Skynet.Throw(Me, "Public Function BuildWith").InvalidInput("Element is nothing")
        
        If TypeOf GenericValues(i) Is IGeneric Then _
            Set Values(i + MIN_INDEX) = GenericValues(i)
      
    Next
    
    Set BuildWith = GenericArray.BuildFrom(Values)
    
End Function

Public Function SetValue(ByVal Value As IGeneric, ParamArray Indices() As Variant)
    
    Dim i As Long
    
    For i = 0 To myClass.Struct.cDims - 1
        myClass.Indices(i) = Indices(i)
    Next
    
    If (SafeArrayPutElement(ByVal myClass.Descriptor, myClass.Indices(0), ByVal ObjPtr(Value)) = DISP_E_BADINDEX) Then _
        Call Skynet.Throw(Me, "Set Object").InvalidInput("Invalid Index")

End Function

Public Function GetValue(ParamArray Indices() As Variant) As IGeneric
    
    Dim i As Long

    Do
        myClass.Indices(i) = Indices(i)
    Loop While InterlockedIncrement(i) < myClass.Struct.cDims

    If (SafeArrayGetElement(ByVal myClass.Descriptor, myClass.Indices(0), GetValue) = DISP_E_BADINDEX) Then _
        Call Skynet.Throw(Me, "Get Object").InvalidInput("Invalid Index")
        
End Function
Public Function ExtractElement(ByVal Descriptor As LongPtr, ByRef Indices() As Long, ByVal Out As LongPtr) As Boolean: ExtractElement = S_OK = SafeArrayGetElement(ByVal Descriptor, Indices(0), ByVal Out): End Function
Public Function AdressOf(ByVal Element As Long) As LongPtr: AdressOf = myClass.Struct.pvData + (Me.SizeOf(Element - MIN_INDEX)): End Function
Public Function SizeOf(ByVal Elements As Long) As Long: SizeOf = Elements * myClass.Struct.cbElements: End Function

Public Function AdressOfIndex(ParamArray Indices() As Variant) As LongPtr
    
    Dim i As Long

    ReDim Index(0 To UBound(Indices)) As Long
    With myClass
        For i = 0 To UBound(Indices)
            Index(i) = Indices(i)
        Next
    End With
    
    Call Catch(SafeArrayPtrOfIndex(ByVal Descriptor, Index(0), AdressOfIndex))
    
End Function

Private Sub SwitchTo(ByVal Descriptor As LongPtr, ByRef s As SAFEARRAYSTRUCT)
    With Skynet
        Call .CopyData(ByVal Descriptor, VarPtr(s), 16&): _
        Call .CopyData(ByVal Descriptor + 16&, VarPtr(s.Bounds(LBound(s.Bounds))), s.cDims * 8&)
    End With
End Sub

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Virtual.Size = LenB(myClass)
End Sub

Private Sub Class_Terminate()
    Call Skynet.Dispose(Me)
End Sub

Public Sub Clear()
    Call Catch(SafeArrayDestroyData(ByVal Me.Descriptor))
End Sub

Public Function SlizeColumn(ByVal Index As Long) As GenericArray
    
    Dim Overlay As SAFEARRAYSTRUCT
    Dim Column As GenericArray
    
    If (Me.Rank <> 2) Then _
        Call Skynet.Throw(Me, "Public Function SlizeColumn").NotSupported("Me.Rank != 2")
            
    If (Index < 1) Or (Index > Me.LengthOf(2)) Then _
        Call Skynet.Throw(Me, "Public Function SlizeColumn").IndexOutofRange("Index not valid")
    
    If (Me.IsEmpty) Then _
        Call Skynet.Throw(Me, "Public Function SlizeColumn").NullPointer("Me.IsEmpty")
        
    Set Column = GenericArray.Build(Me.LengthOf(1))
  
    Call LoadArrayStructure(Overlay, Column.Descriptor)
    Overlay.pvData = Me.AdressOf(Element:=Me.LengthOf(1) * Index)
    
    Call SwitchTo(Me.Descriptor, Overlay)
    Call Catch(SafeArrayCopyData(ByVal Me.Descriptor, ByVal Column.Descriptor))
    Call SwitchTo(Me.Descriptor, myClass.Struct)

    
    Set SlizeColumn = Column
    
End Function

Public Sub Copy( _
                    ByVal SrcArray As GenericArray, _
                    ByVal SrcIndex As Long, _
                    ByVal DstArray As GenericArray, _
                    ByVal DstIndex As Long, _
                    ByVal Elements As Long)
                    
    Dim Overlay As SAFEARRAYSTRUCT
    Dim src As SAFEARRAYSTRUCT
    Dim dst As SAFEARRAYSTRUCT
 
    If (SrcArray Is Nothing) Or (DstArray Is Nothing) Then _
        Call Skynet.Throw(Me, "Public Sub Copy").InvalidInput("SrcArray Is Nothing) Or (DstArray Is Nothing)")
        
    If (Elements < 1) Then _
        Call Skynet.Throw(Me, "Public Sub Copy").InvalidInput("Elements to copy  <  1")
    
    If (SrcIndex < MIN_INDEX) Or (Elements > SrcArray.Length - SrcIndex + MIN_INDEX) Then _
        Call Skynet.Throw(Me, "Public Sub Copy").InvalidInput("(SrcIndex < MIN_INDEX) Or (Elements > SrcArray.Length - SrcIndex + MIN_INDEX)")
        
    If (DstIndex < MIN_INDEX) Or (Elements > DstArray.Length - DstIndex + MIN_INDEX) Then _
        Call Skynet.Throw(Me, "Public Sub Copy").InvalidInput("(DstIndex < MIN_INDEX) Or (Elements > DstArray.Length - DstIndex + MIN_INDEX)")
    
    With Overlay
        .cDims = 1
        .fFeatures = 1088
        .cbElements = 4&
        ReDim .Bounds(1 To 1)
        .Bounds(1).cElements = Elements
        .Bounds(1).lLbound = MIN_INDEX
    End With
    
    With SrcArray
        Call LoadArrayStructure(src, .Descriptor)
        Overlay.pvData = .AdressOf(Element:=SrcIndex)
        Call SwitchTo(.Descriptor, Overlay)
    End With
    
    With DstArray
        Call LoadArrayStructure(dst, .Descriptor)
        Overlay.pvData = .AdressOf(Element:=DstIndex)
        Call SwitchTo(.Descriptor, Overlay)
    End With
    
    Call Catch(SafeArrayCopyData(ByVal SrcArray.Descriptor, ByVal DstArray.Descriptor))
    
    Call SwitchTo(SrcArray.Descriptor, src)
    Call SwitchTo(DstArray.Descriptor, dst)

End Sub

Public Sub ResizeTo(ByVal newSize As Long)
    
    Dim bound As SAFEARRAYBOUND
    
    If Not (Me.Rank = 1) Then _
        Call Skynet.Throw(Me, "Public Sub ReSizeTo").NotSupported("Me.Rank != 1")

''The hard way ;)
'    With myClass.struct
'        .Bounds(1).cElements = newSize
'        Dim newStruct As SAFEARRAYSTRUCT
'        Dim newDescriptor As Long
'        newDescriptor = SafeArrayCreateEx(VbVarType.vbObject, .cDims, .Bounds(1), ByVal 0&)
'        newStruct = GetArrayStructure(newDescriptor)
'        Call Skynet.CopyData(Dst:=newStruct.pvData, src:=.pvData, Bytes:=myClass.Size, ClearSource:=True)
'    End With
'
'    With myClass
'        Call Catch(SafeArrayDestroy(ByVal Descriptor))
'        .Descriptor = newDescriptor
'        .struct = newStruct
'        .Elements = newSize
'        .Size = newSize * .struct.cbElements
'    End With
'
    bound.cElements = newSize
    bound.lLbound = MIN_INDEX
   
    With myClass
         Call Catch(SafeArrayRedim(ByVal .Descriptor, bound))
         Call LoadArrayStructure(.Struct, .Descriptor)
        .Elements = newSize
    End With
    
End Sub
Public Sub Swap(ByVal i As Long, j As Long)
    Dim tmp As IGeneric
    With myClass
        Set tmp = .Internal(i)
        Set .Internal(i) = .Internal(j)
        Set .Internal(j) = tmp
    End With
End Sub
'Public Sub Heap()
'
'    Dim i As Long
'    Dim temp As IGeneric
'
'    Dim hi As Long
'    hi = UBound(myClass.Internal)
'
'    Dim lo As Long
'    lo = LBound(myClass.Internal)
'
'    Dim n As Long
'    n = hi - lo + 1
'
'    Dim start As Long
'    start = (n - 2) \ 2
'
'    Dim eend As Long
'    eend = n - 1
'
'    For i = (n / 2) - 1 To 1 Step -1
'
'        SiftDown List:=myClass.Internal, root:=i, eend:=n - 1, lo:=lo
'    Next
'
'    For i = n - 1 To 1 Step -1
'        Set temp = myClass.Internal(lo + i)
'        Set myClass.Internal(lo + i) = myClass.Internal(lo)
'        Set myClass.Internal(lo) = temp
'        SiftDown List:=myClass.Internal, root:=0, eend:=i - 1, lo:=lo
'    Next
'
'
''    While start >= 0
''        SiftDown myClass.Internal, start, eend
''        start = start - 1
''    Wend
'
'
''    While eend > 0
''        Set temp = myClass.Internal(lo + eend)
''        Set myClass.Internal(lo + eend) = myClass.Internal(lo)
''        Set myClass.Internal(lo) = temp
''
''        eend = eend - 1
''
''        SiftDown myClass.Internal, 0, eend
''    Wend
''
'    For i = 1 To 100
'        If Not myClass.Internal(i) Is Nothing Then
'            Debug.Print i & "//" & myClass.Internal(i)
'        End If
'    Next
'
''    Dim n As Long: n = hi - lo + 1
''    Dim i As Long
''
''    For i = n / 2 To 1 Step -1
''        Call DownHeap(Keys, i, n, lo)
''    Next
''    For i = n To 2 Step -1
''        Call Swap(Keys, lo, lo + i - 1)
''        Call DownHeap(Keys, 1, i - 1, lo)
''    Next
'
'End Sub

'Sub SiftDown(ByRef List() As IGeneric, ByVal root As Long, eend As Long, ByVal lo As Long)
'
'    Dim temp As IGeneric
'    Dim child As Long
'
'    Dim d As IGeneric: Set d = List(lo + root)
'
'    Do While root * 2 + 1 <= eend
'        child = root * 2 + 1
'        If child + 1 <= eend Then
'            If Not List(lo + child) Is Nothing Then
'                If List(lo + child).CompareTo(List(lo + child + 1)) = IsGreater Then
'                    child = child + 1
'                End If
'            End If
'        End If
'
'        If Not List(lo + root) Is Nothing Then
'            If List(lo + root).CompareTo(List(lo + child)) = IsGreater Then
'                Set temp = List(lo + root)
'                Set List(lo + root) = List(lo + child)
'                Set List(lo + child) = temp
'            End If
'        Else
'            Exit Sub
'        End If
'        root = child
'    Loop
'
'End Sub

Public Function BinarySearch( _
                                ByVal Value As IGeneric, _
                                ByVal Index As Long, _
                                ByVal Length As Long, _
                                ByVal Order As SortOrder) As Long

    If Not Me.Rank = 1 Then _
        Call Skynet.Throw(Me, "Public Function BinarySearch").NotSupported("Me.Rank != 1")

    If (Index < MIN_INDEX Or Length < 0) Then _
        Call Skynet.Throw(Me, "Public Function BinarySearch").ArgOutOfRange("Index < MIN_INDEX Or Length < 0")

    If (myClass.Elements - (Index - MIN_INDEX) < Length) Then _
        Call Skynet.Throw(Me, "Public Function BinarySearch").ArgOutOfRange("Argument_InvalidOffLen")

    BinarySearch = Skynet.BinarySearch(Objects:=myClass.Internal, Value:=Value, Index:=Index, Length:=Length, Order:=Order)
   
End Function

Public Function IndexOf( _
                            ByVal Value As IGeneric, _
                            Optional ByVal StartIndex As Long = MIN_INDEX, _
                            Optional ByVal Count As Long) As Long
 
   
    If Not (Me.Rank = 1) Then _
        Call Skynet.Throw(Me, "Public Function IndexOf").NotSupported("Me.Rank != 1")
        
    If (StartIndex < MIN_INDEX) Or (StartIndex > myClass.Elements + MIN_INDEX) Then _
        Call Skynet.Throw(Me, "Public Function IndexOf").IndexOutofRange("(StartIndex < MIN_INDEX) Or (StartIndex > Me.Count() + MIN_INDEX)")

    If (Count = 0) Then
        Count = (myClass.Elements - StartIndex) + MIN_INDEX
    Else
        If (Count < 0) Or (Count > myClass.Elements - StartIndex + MIN_INDEX) Then _
            Call Skynet.Throw(Me, "Public Function IndexOf").ArgOutOfRange("(Count < 0) Or (Count > Me.Elements - StartIndex + MIN_INDEX)")
    End If
   
    IndexOf = Skynet.IndexOf(Objects:=myClass.Internal, Value:=Value, StartIndex:=StartIndex, Count:=Count)
    
End Function

Public Sub Reverse(Optional ByVal Index As Long = MIN_INDEX, Optional ByVal Length As Long)
    
    If Length = 0 Then _
        Length = Me.Length
    
    If Not Me.Rank = 1 Then _
        Call Skynet.Throw(Me, "Public Sub Reverse").NotSupported("Me.Rank != 1")
        
    If (Index < MIN_INDEX Or Length < 0) Then _
        Call Skynet.Throw(Me, "Public Sub Reverse").ArgOutOfRange("Index < MIN_INDEX Or Length < 0")
    
    If (myClass.Elements - (Index - MIN_INDEX) < Length) Then _
        Call Skynet.Throw(Me, "Public Sub Reverse").ArgOutOfRange("Argument_InvalidOffLen")

    Call Skynet.Reverse(Objects:=myClass.Internal, Index:=Index, Length:=Length)
    
End Sub

Public Sub Sort(ByVal Order As SortOrder, Optional ByVal Index As Long = MIN_INDEX, Optional ByVal Length As Long)
    
    If Length = 0 Then _
        Length = myClass.Elements
    
    If Not Me.Rank = 1 Then _
        Call Skynet.Throw(Me, "Public Sub Sort").NotSupported("Me.Rank != 1")
        
    If (Index < MIN_INDEX Or Length < 0) Then _
        Call Skynet.Throw(Me, "Public Sub Sort").ArgOutOfRange("Index < MIN_INDEX Or Length < 0")
    
    If (myClass.Elements - (Index - MIN_INDEX) < Length) Then _
        Call Skynet.Throw(Me, "Public Sub Sort").ArgOutOfRange("Argument_InvalidOffLen")
        
    Call Skynet.Sort(myClass.Internal, Index, Length, Order)

End Sub

Public Sub Transpose()
    
    Const ZEROPOINTER As Long = 0
    
    Dim ptrT1 As LongPtr, ptrT2 As LongPtr
    Dim t1() As ObjectTransfer, t2() As ObjectTransfer
    Dim TotalSize As Long
    Dim i As Long, j As Long, m As Long, n As Long
    
    If (Me.IsEmpty) Then _
        Call Skynet.Throw(Me, "Public Sub Transpose").NullPointer("SafeArray not allocated")
        
    If (Me.Rank <> 2) Then _
        Call Skynet.Throw(Me, "Public Sub Transpose").InvalidOperation("Rank <> 2")
        
    With myClass.Struct
        m = .Bounds(2).cElements
        n = .Bounds(1).cElements
        
        ReDim t1(1 To m, 1 To n)
        ReDim t2(1 To n, 1 To m)

        ptrT1 = VarPtr(t1(1, 1))
        ptrT2 = VarPtr(t2(1, 1))
        TotalSize = Me.SizeOf(Me.Length)
        
        Call Skynet.CopyData(dst:=ByVal ptrT1, src:=ByVal .pvData, Bytes:=TotalSize)
        
        For i = LBound(t1, 1) To UBound(t1, 1)
            For j = LBound(t1, 2) To UBound(t1, 2)
                t2(j, i) = t1(i, j)
            Next
        Next
    
        Call Skynet.CopyData(dst:=ByVal .pvData, src:=ByVal ptrT2, Bytes:=TotalSize)

        Dim Bounds() As SAFEARRAYBOUND
        ReDim Bounds(1 To 2)
        Bounds(1) = .Bounds(2): Bounds(2) = .Bounds(1)

        Call Skynet.CopyData(dst:=ByVal Me.Descriptor + 16&, src:=VarPtr(Bounds(1)), Bytes:=.cDims * 8&)
        .Bounds(1) = Bounds(2): .Bounds(2) = Bounds(1)
     
    End With

'    With Skynet
'        Call .CopyData(ByVal ptrT1, ByVal VarPtr(ZEROPOINTER), .PointerLength)
'        Call .CopyData(ByVal ptrT2, ByVal VarPtr(ZEROPOINTER), .PointerLength)
'    End With
'
End Sub

'Public Sub Feature()
'
'    Const FADF_AUTO = &H1
'    Const FADF_STATIC = &H2
'    Const FADF_EMBEDDED = &H4
'    Const FADF_FIXEDSize = &H10
'    Const FADF_RECORD = &H20
'    Const FADF_HAVEIID = &H40
'    Const FADF_HAVEVARTYPE = &H80
'    Const FADF_BSTR = &H100
'    Const FADF_UNKNOWN = &H200
'    Const FADF_DISPATCH = &H400
'    Const FADF_VARIANT = &H800
'
'
'    Debug.Print myClass.Struct.fFeatures
'
'End Sub

Private Sub Catch(ByVal hResult As Long)

    If Not (hResult = S_OK) Then
        Select Case hResult
            Case E_INVALIDARG
                Debug.Print "One of the arguments is not valid."
            Case DISP_E_BADINDEX
                Debug.Print "The specified index is not valid."
            Case E_OUTOFMEMORY
                Debug.Print "Failed to allocate necessary memory."
            Case DISP_E_ARRAYISLOCKED
                Debug.Print "The array is locked."
            Case E_UNEXPECTED
                Debug.Print "The array could not be locked/ unlocked."
            
            Case Else
                Debug.Print "some error"
        End Select
    End If
    
End Sub

'Todo new instance
Public Function AsReadOnly(Optional ByVal Index As Long = MIN_INDEX, Optional ByVal Count As Long) As IGenericReadOnlyList
  
    If Count = 0 Then _
        Count = myClass.Elements
    
    If Not Me.Rank = 1 Then _
        Call Skynet.Throw(Me, "Public Function AsReadOnly").NotSupported("Me.Rank != 1")
        
    If (Index < MIN_INDEX Or Count < 0) Then _
        Call Skynet.Throw(Me, "Public Function AsReadOnly").ArgOutOfRange("Index < MIN_INDEX Or Count < 0")
    
    If (myClass.Elements - (Index - MIN_INDEX) < Count) Then _
        Call Skynet.Throw(Me, "Public Function AsReadOnly").ArgOutOfRange("Argument_InvalidOffLen")
    
    If (Me.IsEmpty) Then _
        Call Skynet.Throw(Me, "Public Function AsReadOnly").NullPointer("Me.IsEmpty")
    
    Call LoadIterator(myClass.Iterator, Nothing, myClass.Descriptor, myClass.Struct.cDims, Index, Count)
    myClass.IsReadOnly = True
    
    Set AsReadOnly = Me

End Function

Public Function Iterator(Optional ByVal StartIndex As Long = MIN_INDEX, Optional ByVal Count As Long) As IGenericIterator

    Dim c As Member
        
    If Count = 0 Then _
        Count = myClass.Elements
        
    If (myClass.Descriptor = 0) Then _
        Call Skynet.Throw(Me, "Public Function Iterator").InvalidOperation("Descriptor = 0")
            
    If (StartIndex < MIN_INDEX) Or (StartIndex > myClass.Elements + MIN_INDEX) Then _
        Call Skynet.Throw(Me, "Public Function Iterator").IndexOutofRange("(StartIndex < MIN_INDEX) Or (StartIndex > myClass.Elements + MIN_INDEX)")

    If (Count < 0) Or (Count > myClass.Elements - StartIndex + MIN_INDEX) Then _
        Call Skynet.Throw(Me, "Public Function Iterator").ArgOutOfRange("(Count < 0) Or (Count > myClass.Elements - StartIndex + MIN_INDEX)")

    Call LoadIterator(c.Iterator, Me, myClass.Descriptor, myClass.Struct.cDims, StartIndex, Count)
    
    Set Iterator = Skynet.CreateInstance(New GenericArray, VarPtr(c), LenB(c))
    
End Function

Private Sub LoadIterator( _
                            ByRef e As Iterator, _
                            ByVal Parent As GenericArray, _
                            ByVal Descriptor As Long, _
                            ByVal Rank As Long, _
                            ByVal StartIndex As Long, _
                            ByVal Count As Long)

    With e
        Set .Parent = Parent '! default is nothing, should only be set with Iterator
        Set .Current = Nothing
        .Descriptor = Descriptor
        .Rank = Rank
        .Complete = False
        .StartIndex = StartIndex
        .EndIndex = .StartIndex + Count - MIN_INDEX
        .Count = Count
       
        If (.Rank = 1) Then
            .Index = StartIndex - 1
        Else
            ReDim .Indices(.Rank - 1)
            .ptrIndices = VarPtr(.Indices(0))
        End If
        
        .HasLoaded = True
    End With

End Sub
'IGeneric
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Sub IGeneric_Dispose()

'    If (SafeArrayUnlock(myClass.Descriptor) = S_OK) Then
'        Call Catch(SafeArrayDestroyData(ByVal Me.Descriptor))
'        Call Catch(SafeArrayDestroyDescriptor(ByVal Me.Descriptor))
''    Call Catch(SafeArrayDestroy(ByVal Descriptor))
'    End If
    
    If myClass.IsReadOnly Then _
        Call Skynet.ClearData(Virtual.Pointer, Virtual.Size)
  
    myClass = EmptyStructure
  
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member
    
    If (Me.IsEmpty) Then _
        Call Skynet.Throw(Me, "Private Function IGeneric_Clone").NullPointer("Me.IsEmpty")
        
    With c
        Call Catch(SafeArrayCopy(ByVal myClass.Descriptor, .Descriptor))
'        Call Catch(SafeArrayDestroyDescriptor(ByVal Skynet.PointerOf(.Internal)))
        
        Call Skynet.CopyData(ByVal VarPtrArray(.Internal), ByVal VarPtr(.Descriptor), 4&)
        Call LoadArrayStructure(.Struct, .Descriptor) '! pv.data
       
        ReDim .Indices(.Struct.cDims - 1) '!
        .Elements = myClass.Elements
    End With
    
    Set IGeneric_Clone = Skynet.CreateInstance(New GenericArray, VarPtr(c), LenB(c))
  
End Function

Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean: IGeneric_Equals = Skynet.SequenceCompare(Me, Other) = IsEqual: End Function
Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum: IGeneric_CompareTo = Skynet.SequenceCompare(Me, Other): End Function
Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericArray): End Function
Private Function IGeneric_ToString() As String
    IGeneric_ToString = _
                        "Class: " & TypeName$(Me) & vbNewLine & _
                        "Rank: " & Me.Rank & vbNewLine & _
                        "Elements: " & Me.Length & vbNewLine & _
                        "MemorySize: " & Me.TotalSize & " Byte"
End Function

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_BaseIndex() As Long: IGenericReadOnlyList_BaseIndex = MIN_INDEX: End Property
Private Property Get IGenericReadOnlyList_Count() As Long: IGenericReadOnlyList_Count = myClass.Elements: End Property
Private Property Get IGenericReadOnlyList_Item(ByVal Index As Long) As IGeneric
    
'    If (Index > myClass.Elements) Or Index < MIN_INDEX Then _
'        Call Skynet.Throw(Me, "Private Property Get IGenericReadOnlyList_Item").IndexOutofRange("Index not valid") '!
'
    Set IGenericReadOnlyList_Item = myClass.Internal(Index)
        
End Property
Private Function IGenericReadOnlyList_IndexOf(ByVal Item As IGeneric) As Long: IGenericReadOnlyList_IndexOf = Me.IndexOf(Value:=Item, StartIndex:=myClass.Iterator.StartIndex, Count:=myClass.Iterator.Count): End Function
Private Sub IGenericReadOnlyList_Clear(): Call Me.Clear: End Sub
Private Sub IGenericReadOnlyList_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call GenericArray.Copy(SrcArray:=Me, SrcIndex:=myClass.Iterator.StartIndex, DstArray:=GArray, DstIndex:=ArrayIndex, Elements:=myClass.Iterator.Count): End Sub 'TODO CopyTo as Vector
Private Function IGenericReadOnlyList_Iterator() As IGenericIterator: Set IGenericReadOnlyList_Iterator = Me.Iterator(): End Function

'IGenericCollection
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = True: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = myClass.IsReadOnly: End Property
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Length: End Property
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call GenericArray.Copy(SrcArray:=Me, SrcIndex:=MIN_INDEX, DstArray:=GArray, DstIndex:=ArrayIndex, Elements:=Me.Length): End Sub
Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = Me: End Function
Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean: IGenericCollection_Contains = Me.IndexOf(Value:=Item, StartIndex:=MIN_INDEX, Count:=myClass.Elements) >= 0: End Function
Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Iterator(MIN_INDEX, Me.Length): End Function
'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myClass.Iterator.Current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    With myClass.Iterator
        If (.Rank = 1) Then
            If .Index = .EndIndex Then
                If .HasLoaded = False Then _
                    Call Skynet.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator
            
                Set Out = Nothing
                Set .Current = Nothing
                Exit Function
            End If
            
            IGenericIterator_HasNext = (SafeArrayGetElement(ByVal .Descriptor, InterlockedIncrement(.Index), .Current) = S_OK)
            Set Out = .Current
        Else
            Call Skynet.Throw(Me, "Private Function IGenericIterator_HasNext").NotImplemented("multi dim not implemented yet")
            'Multidim
        End If
    End With
    
End Function

Private Sub IGenericIterator_Reset()
    With myClass.Iterator
        If .HasLoaded = False Then _
            Call Skynet.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
            
        Call LoadIterator(myClass.Iterator, .Parent, .Descriptor, .Rank, .StartIndex, .Count)
    End With
End Sub

'IPrototype
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericArray: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GenericArray: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.Size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property




