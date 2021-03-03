VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericArray
'*
'* Purpose: An array stores a fixed-size sequential collection of elements of the IGeneric type.
'*          It provides methods for creating, manipulating, searching, and sorting arrays, thereby serving as the base class for all Array instances.
'*          An element is a (reference) value in an Array. The length of an Array is the total number of elements it can contain.
'*          The lower bound of an Array is the index of its first element (one by default).
'*          A multidimensional Array can have different upper bounds for each dimension.
'*          An array can have a maximum of 32 dimensions.
'*          The Array class is part of the IGenericCollection namespaces. Unlike the other classes which are the IGenericCollection implemented,
'*          Array has a fixed capacity. To increase the capacity, you must create a new Array object with the required capacity,
'*          copy the elements from the old Array object to the new one, and delete the old Array.
'*          The Array is not guaranteed to be sorted. You must sort the Array prior to performing operations (such as BinarySearch)
'*          that require the Array to be sorted.
'*
'* Notes: An array can be Single-Dimensional, Multidimensional or Jagged.
'*        The number of dimensions and the length of each dimension are established when the array instance is created.
'*        These values can't be changed during the lifetime of the instance.
'*        The default values (reference elements) of an array elements are set to null.
'*        Each element can be referred to by an index.
'*        The length of the array is the total number of item present in the array.
'*        A jagged array is an array of arrays, and therefore its elements are reference types and are initialized to null.
'*        Arrays are one indexed: an array with n elements is indexed from 1 to n.
'*        Array elements can be of any IGeneric type, including an array type.
'*        Array types are reference types. GenericArray implements IGenericIterator, you can use foreach iteration.
'*
'*
'*======================================================================================================================

'@Folder("<T>Enumerable")

Option Explicit
'@PredeclaredId

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
'SafeArray fFeatures
Private Const FADF_AUTO As Long = &H1
Private Const FADF_STATIC As Long = &H2
Private Const FADF_EMBEDDED As Long = &H4
Private Const FADF_FIXEDSize As Long = &H10
Private Const FADF_RECORD As Long = &H20
Private Const FADF_HAVEIID As Long = &H40
Private Const FADF_HAVEVARTYPE As Long = &H80
Private Const FADF_BSTR As Long = &H100
Private Const FADF_UNKNOWN As Long = &H200
Private Const FADF_DISPATCH As Long = &H400
Private Const FADF_VARIANT As Long = &H800
Private Const FADF_CREATEVECTOR As Long = &H2000 ' Flag for safearraydestroydata

Private Const MIN_INDEX As Long = 0 'Option Base { 0 | 1 }
Private Const MAX_ARRAY_LENGHT As Long = &H7FEFFFFF
Private Const F_FEATURES As Long = 1088
Private Const ELEMENT_SIZE As Long = 4&
Private Const ELEMENT_TYPE As Integer = 9 'vbvartype.vbObject

Private Const DEFAULT_SORTORDER As Integer = 1 'SortOrder.Ascending
Private Const QUICKSORT_DEPTH_THRESHOLD As Long = 30 'Should be <=32
Private Const INSERTIONSORT_THRESHOLD As Long = 16 'Sould be small
Private Const COPYMEMORY_THRESHOLD As Long = 128 'Depends on machine

Public Enum SortMethod
    RandomSort = 0
    IntrospectiveSort = -1
    QuickSort = 1
End Enum

#If Win64 Then
    Private Const POINTERSIZE As LongPtr = 8
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dst As Any, ByRef Src As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef Dst As Any, ByVal Length As LongPtr, ByVal Fill As Byte)
    Private Declare PtrSafe Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As LongPtr, ByRef Dst As LongPtr, ByVal Src As LongPtr) As Long
    Private Declare PtrSafe Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Length As Long, ByRef Dst As LongPtr, ByVal Src As LongPtr) As Long
    Private Declare PtrSafe Function SafeArrayCreateEx Lib "oleaut32.dll" (ByVal vt As Integer, ByVal cDims As Long, ByRef rgsabound As SAFEARRAYBOUND, ByRef pvExtra As Any) As LongPtr
    Private Declare PtrSafe Function SafeArrayAllocData Lib "oleaut32.dll" (ByRef psa As Any) As LongPtr
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
    Private Const POINTERSIZE As Long = 4
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dst As Any, ByRef Src As Any, ByVal Length As Long)
    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Dst As Any, ByVal Length As Long)
    Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (ByRef Dst As Any, ByVal Length As Long, ByVal Fill As Byte)
    Private Declare Function CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal Length As Long, ByRef Dst As Long, ByVal Src As Long) As Long
    Private Declare Function CopyBytesZero Lib "msvbvm60.dll" Alias "__vbaCopyByteZeros" (ByVal Length As Long, ByRef Dst As Long, ByVal Src As Long) As Long
    Private Declare Function SafeArrayCreateEx Lib "oleaut32.dll" (ByVal vt As Integer, ByVal cDims As Long, ByRef rgsabound As SAFEARRAYBOUND, ByRef pvExtra As Any) As Long
    Private Declare Function SafeArrayAllocData Lib "oleaut32.dll" (ByRef psa As Any) As Long
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
    #If Win64 Then
        pvData       As LongPtr
    #Else
        pvData       As Long
    #End If
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
    HasLoaded As Boolean
    Index As Long
    EndIndex As Long
    current As IGeneric
    Matrix As GenericArray
    Rank As Long
End Type

Private Type Member
   Struct As SAFEARRAYSTRUCT
   Descriptor As LongPtr
   Elements As Long
   Indices() As Long
   Internal() As IGeneric
   Iterator As Iterator
End Type
Private myclass As Member

Public Property Get Elements() As IGenericCollection: Set Elements = Me: End Property
Public Property Get MaxLength() As Long: MaxLength = MAX_ARRAY_LENGHT: End Property
Public Property Get LowerBound() As Long: LowerBound = MIN_INDEX: End Property
Public Property Get IsEmpty() As Boolean: IsEmpty = (myclass.Elements = 0): End Property
Public Property Get Descriptor() As LongPtr: Descriptor = myclass.Descriptor: End Property
'Public Property Get LengthOf(ByVal Dimension As Long) As Long: Call Catch(SafeArrayGetUBound(ByVal myclass.Descriptor, Dimension, LengthOf)): End Property
Public Property Get LengthOf(ByVal Dimension As Long) As Long: LengthOf = myclass.Struct.Bounds(myclass.Struct.cDims - Dimension).cElements: End Property
Public Property Get Length() As Long: Length = myclass.Elements: End Property
Public Property Get Size() As LongPtr: Size = Me.SizeOf(myclass.Elements): End Property
'Public Property Get Rank() As Long: Rank = SafeArrayGetDim(ByVal myclass.Descriptor): End Property
Public Property Get Rank() As Long: Rank = myclass.Struct.cDims: End Property
Public Property Set ElementAt(ByVal Index As Long, ByVal Value As IGeneric): Set myclass.Internal(Index) = Value: End Property
Public Property Get ElementAt(ByVal Index As Long) As IGeneric: Set ElementAt = myclass.Internal(Index): End Property

Public Function Build(ParamArray Indices() As Variant) As GenericArray
    
    Dim newArray As Member
    Dim SAB() As SAFEARRAYBOUND
    Dim i As Long, UB As Long, LB As Long
                                                                                            
    If IsMissing(Indices) Then _
        Call System.Throw(Me, "Public Function Build").InvalidInput("ParamArray Indices are missing")
    
    LB = 0&: UB = UBound(Indices)
    
    ReDim newArray.Indices(LB To UB)
    ReDim SAB(LB To UB)
    
    newArray.Elements = 1
    newArray.Struct.cDims = UB + 1
    newArray.Struct.fFeatures = F_FEATURES
    newArray.Struct.cbElements = ELEMENT_SIZE
    ReDim newArray.Struct.Bounds(LB To UB)
        
    For i = LB To UB
        If Not (Indices(i) >= 0) Then _
            Call System.Throw(Me, "Build").ArgOutOfRange("A zero(one)-based array indexed array has its first item indexed as 0(1) and need to have at least one element")

        SAB(i).cElements = Indices(i)
        SAB(i).lLbound = MIN_INDEX

        newArray.Elements = newArray.Elements * Indices(i)
        newArray.Indices(i) = MIN_INDEX
        newArray.Struct.Bounds(UB - i) = SAB(i)
    Next i

    newArray.Descriptor = SafeArrayCreateEx(ELEMENT_TYPE, newArray.Struct.cDims, SAB(LB), ByVal 0&)

    If (newArray.Descriptor = 0) Then _
        Call System.Throw(Me, "Public Function Build").InvalidOperation("Descriptor = 0")
    Call LoadArray(newArray.Internal, newArray.Descriptor)
    Call Catch(SafeArrayPtrOfIndex(ByVal newArray.Descriptor, newArray.Indices(LB), newArray.Struct.pvData))
 
    Set Build = System.CreateInstance(New GenericArray, VarPtr(newArray), LenB(newArray))
    
End Function

Public Function AsArray(ParamArray GenericValues() As Variant) As GenericArray
    
    Dim i As Long
    Dim Values As GenericArray
    
    If IsMissing(GenericValues) Then _
        Call System.Throw(Me, "Public Function Create").Instantiation("ParamArray GenericValues are missing")
    
    Set Values = GenericArray.Build(UBound(GenericValues) + 1)
    
    With Values
        For i = LBound(GenericValues) To UBound(GenericValues)
            
            If Not VarType(GenericValues(i)) = ELEMENT_TYPE Then _
                Call System.Throw(Me, "Public Function Create").InvalidInput("Datatype {" & TypeName$(GenericValues(i)) & "} not allowed")
            
            If (GenericValues(i) Is Nothing) Then _
                Call System.Throw(Me, "Public Function Create").InvalidInput("Element is nothing")
            
            If Not TypeOf GenericValues(i) Is IGeneric Then _
                Call System.Throw(Me, "Public Function Create").InvalidInput("Not TypeOf GenericValues(i) Is IGeneric")
                
            Set .ElementAt(i) = GenericValues(i)
          
        Next
    End With
    
    Set AsArray = Values
    
End Function
'' Works and very fast for large matrices
'Public Function BuildFrom(ByRef Matrix() As IGeneric) As GenericArray
'
'    Dim i As Long, FormatBounds As Boolean
'    Dim newArray As Member
'
'    If Not VarType(Matrix) = VbVarType.vbArray + ELEMENT_TYPE Then _
'        Call System.Throw(Me, "Public Function BuildWith").InvalidInput("Datatype {" & TypeName$(Matrix) & "} not allowed")
'
'    Call CopyMemory(Dst:=newArray.Descriptor, Src:=ByVal VarPtrArray(Matrix), Length:=POINTERSIZE)
'
'    If (newArray.Descriptor = 0) Then _
'        Call System.Throw(Me, "Public Function BuildFrom").InvalidOperation("Descriptor = 0")
'    Call ZeroMemory(Dst:=ByVal VarPtrArray(Matrix), Length:=POINTERSIZE) 'FillMemory(dst:=ByVal VarPtrArray(Matrix), Length:=POINTERSIZE, Fill:=0)
'
'    newArray.Elements = 1
'    Call LoadArrayStructure(newArray.Struct, newArray.Descriptor)
'
'    ReDim newArray.Indices(newArray.Struct.cDims - 1)
'    For i = LBound(newArray.Struct.Bounds) To UBound(newArray.Struct.Bounds)
'        If Not newArray.Struct.Bounds(i).lLbound = MIN_INDEX Then _
'            FormatBounds = True
'
'        newArray.Struct.Bounds(i).lLbound = MIN_INDEX
'        newArray.Elements = newArray.Elements * newArray.Struct.Bounds(i).cElements
'    Next
'
'    If FormatBounds Then _
'        Call CopyMemory(Dst:=ByVal myclass.Descriptor + 16&, Src:=ByVal VarPtr(newArray.Struct.Bounds(0&)), Length:=newArray.Struct.cDims * 8&)
'
'    Call LoadArray(newArray.Internal, newArray.Descriptor)
'
'    Set BuildFrom = System.CreateInstance(New GenericArray, VarPtr(newArray), LenB(newArray))
'
'End Function

Private Sub LoadArrayStructure(ByRef structure As SAFEARRAYSTRUCT, ByVal SafeArrayPointer As LongPtr)
    
    Dim i As Long
    
    If (SafeArrayPointer = 0) Then _
        Call System.Throw(Me, "Private Sub LoadArrayStructure").InvalidInput("SafeArrayPointer")

    Call CopyMemory(Dst:=structure, Src:=ByVal SafeArrayPointer, Length:=12& + POINTERSIZE)
    
    With structure
        If (.cDims <= 0) Then _
            Call System.Throw(Me, "Private Sub LoadArrayStructure").Overflow(".cDims <= 0")
        
        ReDim .Bounds(.cDims - 1)
        Call CopyMemory(Dst:=.Bounds(0&), Src:=ByVal (SafeArrayPointer + 16&), Length:=.cDims * 8&)
    End With

End Sub

Private Sub LoadArray(ByRef Internal() As IGeneric, ByVal Descriptor As LongPtr)
    Call CopyMemory(Dst:=ByVal VarPtrArray(Internal), Src:=ByVal VarPtr(Descriptor), Length:=POINTERSIZE)
End Sub

Public Function DescriptorOf(ByRef SomeArray As Variant) As LongPtr
    
    Const VT_BYREF As Long = &H4000
    Dim vType As Long
    
    If Not (VarType(SomeArray) And VbVarType.vbArray) = VbVarType.vbArray Then _
        System.Throw(Me, "ArrayPointer").InvalidInput ("Not (IsArray(SomeArray))")

    Call CopyMemory(Dst:=vType, Src:=ByVal VarPtr(SomeArray), Length:=2&)
    Call CopyMemory(Dst:=DescriptorOf, Src:=ByVal VarPtr(SomeArray) + 8&, Length:=POINTERSIZE)
    
    'if SomeArray is defined x() as ...
    If (vType And VT_BYREF) = VT_BYREF Then _
        Call CopyMemory(Dst:=DescriptorOf, Src:=ByVal DescriptorOf, Length:=POINTERSIZE)
   
End Function

Public Function SetValue(ByVal Value As IGeneric, ParamArray Indices() As Variant)

    Dim i As Long

    For i = 0 To myclass.Struct.cDims - 1
        myclass.Indices(i) = -1
        myclass.Indices(i) = Indices(i)
    Next

    If (SafeArrayPutElement(ByVal myclass.Descriptor, myclass.Indices(0), ByVal ObjPtr(Value)) = DISP_E_BADINDEX) Then _
        Call System.Throw(Me, "Set Object").InvalidInput("Invalid Index")

End Function

Public Function GetValue(ParamArray Indices() As Variant) As IGeneric
    
    Dim i As Long

    For i = 0 To myclass.Struct.cDims - 1
        myclass.Indices(i) = -1
        myclass.Indices(i) = Indices(i)
    Next

    If (SafeArrayGetElement(ByVal myclass.Descriptor, myclass.Indices(0), GetValue) = DISP_E_BADINDEX) Then _
        Call System.Throw(Me, "Get Object").InvalidInput("Invalid Index")
        
End Function
Public Function ExtractElement(ByVal Descriptor As LongPtr, ByRef Indices() As Long, ByVal Out As LongPtr) As Boolean: ExtractElement = S_OK = SafeArrayGetElement(ByVal Descriptor, Indices(0), ByVal Out): End Function
Public Function AdressOf(Optional ByVal Index As Long = MIN_INDEX) As LongPtr: AdressOf = myclass.Struct.pvData + Me.SizeOf(Index - MIN_INDEX): End Function
Public Function SizeOf(ByVal Elements As Long) As Long: SizeOf = Elements * myclass.Struct.cbElements: End Function

Public Function AdressOfIndex(ParamArray Indices() As Variant) As LongPtr
    
    Dim i As Long

    ReDim Index(0 To UBound(Indices)) As Long
    With myclass
        For i = 0 To UBound(Indices)
            Index(i) = Indices(i)
        Next
    End With

    Call Catch(SafeArrayPtrOfIndex(ByVal Descriptor, Index(0), AdressOfIndex))

End Function

Private Sub SwitchTo(ByVal Descriptor As LongPtr, ByRef s As SAFEARRAYSTRUCT)
    Call CopyMemory(Dst:=ByVal Descriptor, Src:=ByVal VarPtr(s), Length:=16&)
    Call CopyMemory(Dst:=ByVal Descriptor + 16&, Src:=ByVal VarPtr(s.Bounds(LBound(s.Bounds))), Length:=s.cDims * 8&)
End Sub

Private Sub Class_Terminate()
    myclass = EmptyStructure
    'Call Catch(SafeArrayDestroy(ByVal .Descriptor))
End Sub

Public Function SlizeColumn(ByVal Index As Long) As GenericArray
    
    Dim Overlay As SAFEARRAYSTRUCT
    Dim Column As GenericArray
    
    If (Me.Rank <> 2) Then _
        Call System.Throw(Me, "Public Function SlizeColumn").NotSupported("Me.Rank != 2")
            
    If (Index < MIN_INDEX) Or (Index > Me.LengthOf(2)) Then _
        Call System.Throw(Me, "Public Function SlizeColumn").IndexOutofRange("Index not valid")
    
    If (Me.IsEmpty) Then _
        Call System.Throw(Me, "Public Function SlizeColumn").InvalidOperation("Me.IsEmpty")
        
    Set Column = GenericArray.Build(Me.LengthOf(1))
   
    Overlay.cDims = 1
    Overlay.fFeatures = F_FEATURES
    Overlay.cbElements = ELEMENT_SIZE
    ReDim Overlay.Bounds(0)
    Overlay.Bounds(0).cElements = Me.LengthOf(1)
    Overlay.Bounds(0).lLbound = MIN_INDEX
    Overlay.pvData = Me.AdressOf(Index:=(Me.LengthOf(1) * Index))
    
    Call SwitchTo(myclass.Descriptor, Overlay)
    Call Catch(SafeArrayCopyData(ByVal myclass.Descriptor, ByVal Column.Descriptor))
    Call SwitchTo(myclass.Descriptor, myclass.Struct)

    Set SlizeColumn = Column
    
End Function

Public Sub CopyTo( _
                    ByVal GArray As GenericArray, _
                    ByVal ArrayIndex As Long, _
                    Optional ByVal Index As Long = MIN_INDEX, _
                    Optional ByVal Length As Long = -1)

    Dim i As Long
   
    If (GArray Is Nothing) Then _
        Call System.Throw(Me, "Public Sub CopyTo").InvalidInput("(GArray Is Nothing)")
            
    If (myclass.Elements > GArray.Length - ArrayIndex) Then _
        Call System.Throw(Me, "Public Sub CopyTo").InvalidInput("(myclass.Elements > GArray.Length - ArrayIndex")
    
    If myclass.Elements = 0 Then _
        Exit Sub
    
    If Length = -1 Then _
        Length = myclass.Elements
        
    If myclass.Struct.cDims = 1 And GArray.Rank = myclass.Struct.cDims Then
        If myclass.Elements <= COPYMEMORY_THRESHOLD Then
            With GArray
                For i = Index To Length - 1
                    Set .ElementAt(ArrayIndex) = myclass.Internal(i)
                    ArrayIndex = ArrayIndex + 1
                Next
            End With
       End If
    
    Else
    
        Call CopyArrays(Me, Index, GArray, ArrayIndex, Length)
  
    End If

End Sub

Private Sub CopyArrays( _
                        ByVal SrcArray As GenericArray, _
                        ByVal SrcIndex As Long, _
                        ByVal DstArray As GenericArray, _
                        ByVal DstIndex As Long, _
                        ByVal Elements As Long)

    Dim i As Long
    Dim Overlay As SAFEARRAYSTRUCT
    Dim Src As SAFEARRAYSTRUCT
    Dim Dst As SAFEARRAYSTRUCT

    If (SrcArray Is Nothing) Or (DstArray Is Nothing) Then _
        Call System.Throw(Me, "Public Sub CopyArrays").InvalidInput("SrcArray Is Nothing) Or (DstArray Is Nothing)")

    If (Elements < 1) Then _
        Call System.Throw(Me, "Public Sub CopyArrays").InvalidInput("Elements to copy  <  1")

    If (SrcIndex < MIN_INDEX) Or (Elements > SrcArray.Length - SrcIndex) Then _
        Call System.Throw(Me, "Public Sub CopyArrays").InvalidInput("(SrcIndex < MIN_INDEX) Or (Elements > SrcArray.Length - SrcIndex)")

    If (DstIndex < MIN_INDEX) Or (Elements > DstArray.Length - DstIndex) Then _
        Call System.Throw(Me, "Public Sub CopyArrays").InvalidInput("(DstIndex < MIN_INDEX) Or (Elements > DstArray.Length - DstIndex)")

    If DstArray.Rank = 1 And SrcArray.Rank = 1 And Elements <= COPYMEMORY_THRESHOLD Then
        With SrcArray
            For i = SrcIndex To SrcIndex + Elements - 1
                Set DstArray.ElementAt(DstIndex) = .ElementAt(i): DstIndex = DstIndex + 1
            Next
        End With
        Exit Sub
    End If

    Overlay.cDims = 1
    Overlay.fFeatures = F_FEATURES
    Overlay.cbElements = ELEMENT_SIZE
    ReDim Overlay.Bounds(0)
    Overlay.Bounds(0).cElements = Elements
    Overlay.Bounds(0).lLbound = MIN_INDEX

    With SrcArray
        Call LoadArrayStructure(Src, .Descriptor)
        Overlay.pvData = .AdressOf(Index:=SrcIndex)
        Call SwitchTo(.Descriptor, Overlay)
    End With

    With DstArray
        Call LoadArrayStructure(Dst, .Descriptor)
        Overlay.pvData = .AdressOf(Index:=DstIndex)
        Call SwitchTo(.Descriptor, Overlay)
    End With

    Call Catch(SafeArrayCopyData(ByVal SrcArray.Descriptor, ByVal DstArray.Descriptor))

    Call SwitchTo(SrcArray.Descriptor, Src)
    Call SwitchTo(DstArray.Descriptor, Dst)

End Sub
'Haram ;) - 4 legacy
'Public Sub ResizeTo(ByVal newSize As Long)
'
'    Dim bound As SAFEARRAYBOUND
'
'    If Not (Me.Rank = 1) Then _
'        Call System.Throw(Me, "Public Sub ReSizeTo").NotSupported("Me.Rank != 1")
'
'''The hard way ;)
''    With myClass.struct
''        .Bounds(1).cElements = newSize
''        Dim newStruct As SAFEARRAYSTRUCT
''        Dim newDescriptor As Long
''        newDescriptor = SafeArrayCreateEx(VbVarType.vbObject, .cDims, .Bounds(1), ByVal 0&)
''        newStruct = GetArrayStructure(newDescriptor)
''        Call Skynet.CopyData(Dst:=newStruct.pvData, src:=.pvData, Bytes:=myClass.Size, ClearSource:=True)
''    End With
''
''    With myClass
''        Call Catch(SafeArrayDestroy(ByVal Descriptor))
''        .Descriptor = newDescriptor
''        .struct = newStruct
''        .Elements = newSize
''        .Size = newSize * .struct.cbElements
''    End With
''
'    bound.cElements = newSize
'    bound.lLbound = MIN_INDEX
'
'    With myclass
'         Call Catch(SafeArrayRedim(ByVal .Descriptor, bound))
'         Call LoadArrayStructure(.Struct, .Descriptor)
'        .Elements = newSize
'    End With
'
'End Sub

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
                                ByVal Order As SortOrder, _
                                Optional ByVal Index As Long = MIN_INDEX, _
                                Optional ByVal Length As Long = -1, _
                                Optional ByVal Comparer As IGenericComparer) As Long
    
    If Length = -1 Then _
        Length = myclass.Elements - Index
        
'    if (me.Rank != 1)
'    if (Index < min_index or Index > me.Length )
'    if (Length < 0 or Length > me.Length - Index)
'    Call Validate("Public Function BinarySearch", Me, Index, Length) 'GenericSortedList.Build().AddAll(anyCollection) throws error
    BinarySearch = Me.BinarySearchArray(Objects:=myclass.Internal, Value:=Value, LowerBound:=Index, UpperBound:=Length + Index - 1, Order:=Order, Comparer:=Comparer)
  
End Function

Public Function IndexOf( _
                            ByVal Value As IGeneric, _
                            Optional ByVal Index As Long = MIN_INDEX, _
                            Optional ByVal Length As Long = -1, _
                            Optional ByVal Comparer As IGenericComparer) As Long
  
    If Length = -1 Then _
        Length = myclass.Elements - Index
        
'    if (me.Rank != 1)
'    if (Index < min_index or Index > me.Length )
'    if (Length < 0 or Length > me.Length - Index)
    
'    Call Validate("Public Function IndexOf", Me, Index, Length)
    IndexOf = Me.IndexOfArray(Objects:=myclass.Internal, Value:=Value, LowerBound:=Index, UpperBound:=Length + Index - 1, Comparer:=Comparer)
    
End Function

Public Function Reverse( _
                            Optional ByVal Index As Long = MIN_INDEX, _
                            Optional ByVal Length As Long = -1) As GenericArray
    
    If Length = -1 Then _
        Length = myclass.Elements - Index
    
'    if (me.Rank != 1)
'    if (Index < min_index or Index > me.Length )
'    if (Length < 0 or Length > me.Length - Index)
'    Call Validate("Public Function Reverse", Me, Index, Length)
    Call GenericArray.ReverseArray(Objects:=myclass.Internal, LowerBound:=Index, UpperBound:=Length + Index - 1)
    
    Set Reverse = Me
    
End Function

Public Function Shuffle(Optional ByVal Index As Long = MIN_INDEX, Optional ByVal Length As Long = -1) As GenericArray
     
    If Length = -1 Then _
        Length = myclass.Elements - Index
        
    Call Validate("Public Function Shuffle", Me, Index, Length)
    
    If Length > 1 Then _
        Call ShuffleArray(Element:=myclass.Internal, LowerBound:=Index, UpperBound:=Length + Index - 1)
    
    Set Shuffle = Me
    
End Function

Public Function Sort( _
                        Optional ByVal Order As SortOrder = DEFAULT_SORTORDER, _
                        Optional ByVal Index As Long = MIN_INDEX, _
                        Optional ByVal Length As Long = -1, _
                        Optional ByVal Comparer As IGenericComparer) As GenericArray
     
    If Length = -1 Then _
        Length = myclass.Elements - Index
        
    Call Validate("Public Function Sort", Me, Index, Length)
    
    If Length > 1 Then _
        Call GenericArray.SortArray(Objects:=myclass.Internal, LowerBound:=Index, UpperBound:=Length + Index - 1, Order:=Order, Comparer:=Comparer)

    Set Sort = Me
    
End Function

Private Sub Validate(ByRef Caller As String, ByVal GArray As GenericArray, ByVal Index As Long, ByVal Count As Long)
    
    If (GArray Is Nothing) Then _
        Call System.Throw(Me, Caller).InvalidInput("(GArray Is Nothing)")
        
    With GArray
        If (.IsEmpty) Then _
            Call System.Throw(Me, Caller).InvalidOperation("(IsEmpty)")
        
        If Not .Rank = 1 Then _
            Call System.Throw(Me, Caller).NotSupported("Rank != 1")
            
        If (Index < MIN_INDEX Or Index > .Length) Then _
            Call System.Throw(Me, Caller).ArgOutOfRange("(Index < MIN_INDEX Or Index > Length)")
            
        If (Count < 0 Or Count > .Length - Index) Then _
            Call System.Throw(Me, Caller).ArgOutOfRange("(Count < 0 Or Count > .Length - Index) ")
    
    End With

End Sub

Public Sub Transpose()
    
    Const ZEROPOINTER As Long = 0
    
    Dim ptrT1 As LongPtr, ptrT2 As LongPtr
    Dim t1() As ObjectTransfer, t2() As ObjectTransfer
    Dim TotalSize As Long
    Dim i As Long, j As Long, m As Long, N As Long
    
    If (Me.IsEmpty) Then _
        Call System.Throw(Me, "Public Sub Transpose").InvalidOperation("SafeArray not allocated")
        
    If (Me.Rank <> 2) Then _
        Call System.Throw(Me, "Public Sub Transpose").InvalidOperation("Rank <> 2")
        
    With myclass.Struct
        m = .Bounds(1).cElements
        N = .Bounds(0).cElements
        
        ReDim t1(1 To m, 1 To N)
        ReDim t2(1 To N, 1 To m)

        ptrT1 = VarPtr(t1(1, 1))
        ptrT2 = VarPtr(t2(1, 1))
        TotalSize = Me.SizeOf(myclass.Elements)
        
        Call CopyMemory(Dst:=ByVal ptrT1, Src:=ByVal .pvData, Length:=TotalSize)
        
        For i = LBound(t1, 1) To UBound(t1, 1)
            For j = LBound(t1, 2) To UBound(t1, 2)
                t2(j, i) = t1(i, j)
            Next
        Next
    
        Call CopyMemory(Dst:=ByVal .pvData, Src:=ByVal ptrT2, Length:=TotalSize)
         
        Dim Bounds() As SAFEARRAYBOUND
        ReDim Bounds(1)
        Bounds(0) = .Bounds(1): Bounds(1) = .Bounds(0)
        
        Call CopyMemory(Dst:=ByVal myclass.Descriptor + 16&, Src:=ByVal VarPtr(Bounds(0)), Length:=.cDims * 8&)
        .Bounds(0) = Bounds(1): .Bounds(1) = Bounds(0)
     
    End With

End Sub

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

Private Sub LoadIterator(ByRef Iterator As Iterator, ByVal Matrix As GenericArray)

     Set Iterator.Matrix = Matrix
     Iterator.Rank = Matrix.Rank
     Iterator.Index = Matrix.LowerBound
     Iterator.EndIndex = Matrix.Length
     
'     If Rank > 1 Then
'        Iterator.Descriptor = Matrix.Descriptor
'        ReDim Iterator.Indices(Iterator.Rank - 1)
'        Iterator.ptrIndices = VarPtr(Iterator.Indices(0))
'     End If
     
     Iterator.HasLoaded = True

End Sub

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

Public Function BinarySearchArray( _
                                    ByRef Objects() As IGeneric, _
                                    ByVal Value As IGeneric, _
                                    ByVal LowerBound As Long, _
                                    ByVal UpperBound As Long, _
                                    ByVal Order As SortOrder, _
                                    Optional ByVal Comparer As IGenericComparer) As Long

    Dim i As Long
    Dim c As CompareResult
    Dim lo As Long: lo = LowerBound
    Dim hi As Long: hi = UpperBound
    
    If Comparer Is Nothing Or Comparer Is IGenericComparer Then
    
        Dim Element As IGenericValue
        
        Do While (lo <= hi)

            i = lo + ((hi - lo) / 2)
            Set Element = Objects(i) 'Try cast
            If (Element Is Nothing) = False Then
                c = Element.CompareTo(Value) * Order
            Else
                c = IIf(Value Is Nothing, CompareResult.IsEqual, CompareResult.IsLower * Order)
            End If

            If (c = IsEqual) Then _
                BinarySearchArray = i: Exit Function

            If (c = IsLower) Then
                lo = i + 1
            Else
                hi = i - 1
            End If
        Loop
    
    Else
        
        With Comparer
            Do While (lo <= hi)
            
                i = lo + ((hi - lo) / 2)
    
                c = .Compare(Objects(i), Value) * Order
    
                If (c = IsEqual) Then _
                    BinarySearchArray = i: Exit Function
                
                If (c = IsLower) Then
                    lo = i + 1
                Else
                    hi = i - 1
                End If
        
            Loop
        End With
    
    End If

    BinarySearchArray = Not lo
    
End Function

'Returns the index of the first occurrence of a given value in a range of
'this list. It does a linear, O(n) search. The list is searched forwards, starting at index
'and upto count number of elements. The
'elements of the list are compared to the given value using the
'Object.Equals method.

Public Function IndexOfArray( _
                                ByRef Objects() As IGeneric, _
                                ByVal Value As IGeneric, _
                                ByVal LowerBound As Long, _
                                ByVal UpperBound As Long, _
                                Optional ByVal Comparer As IGenericComparer) As Long
    
    Dim i As Long
        
    If Not (Value Is Nothing) Then
        If Comparer Is Nothing Or Comparer Is IGenericComparer Then
            
            For i = LowerBound To UpperBound
                If Value.Equals(Objects(i)) Then _
                    IndexOfArray = i: Exit Function
            Next
    
        Else
        
            With Comparer
                For i = LowerBound To UpperBound
                    If .Equals(Objects(i), Value) Then _
                        IndexOfArray = i: Exit Function
                Next
            End With
            
        End If
        
    Else
    
        For i = LowerBound To UpperBound
            If (Objects(i) Is Nothing) Then _
                IndexOfArray = i: Exit Function
        Next
        
    End If
    
    IndexOfArray = -1
    
End Function
'Reverses the elements in a range of this list. Following a call to this
'method, an element in the range given by index and count
'which was previously located at index i will now be located at
'index index + (index + count - i - 1).
Public Sub ReverseArray( _
                            ByRef Objects() As IGeneric, _
                            ByVal LowerBound As Long, _
                            ByVal UpperBound As Long)

    Dim i As Long: i = LowerBound
    Dim j As Long: j = UpperBound
    Dim temp As IGeneric
    
    Do While i < j
        Set temp = Objects(i)
        Set Objects(i) = Objects(j)
        Set Objects(j) = temp
        Call InterlockedIncrement(i)
        Call InterlockedDecrement(j)
    Loop 'While InterlockedIncrement(i) < InterlockedDecrement(j)

End Sub

'Sorting
'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub SortArray( _
                        ByRef Objects() As IGeneric, _
                        ByVal LowerBound As Long, _
                        ByVal UpperBound As Long, _
                        Optional ByVal Comparer As IGenericComparer, _
                        Optional ByVal Order As SortOrder = DEFAULT_SORTORDER)
    
    If Comparer Is Nothing Or Comparer Is IGenericComparer Then
        Call DepthLimitedQuickSort(Objects, LowerBound, UpperBound, QUICKSORT_DEPTH_THRESHOLD, Order)   'Try Convert to GenericValue
    Else
        Call DepthLimitedQuickSort_Comparer(Objects, LowerBound, UpperBound, QUICKSORT_DEPTH_THRESHOLD, Order, Comparer)
    End If

End Sub

Public Sub SortArrayWithItems( _
                                ByRef Keys() As IGeneric, _
                                ByRef Items() As Object, _
                                ByVal LowerBound As Long, _
                                ByVal UpperBound As Long, _
                                Optional ByVal Comparer As IGenericComparer, _
                                Optional ByVal Order As SortOrder = DEFAULT_SORTORDER)
 
    If Comparer Is Nothing Or Comparer Is IGenericComparer Then
        Call DepthLimitedQuickSortWithItems(Keys, Items, LowerBound, UpperBound, QUICKSORT_DEPTH_THRESHOLD, Order)  'Try Convert to GenericValue
    Else
        Call DepthLimitedQuickSortWithItems_Comparer(Keys, Items, LowerBound, UpperBound, QUICKSORT_DEPTH_THRESHOLD, Order, Comparer)
    End If
   
End Sub

Public Sub SortGenericArrays( _
                                ByVal Keys As GenericArray, _
                                ByVal Items As GenericArray, _
                                ByVal Index As Long, _
                                ByVal Length As Long, _
                                Optional ByVal Comparer As IGenericComparer, _
                                Optional ByVal Order As SortOrder = DEFAULT_SORTORDER)

    Dim KeyArray() As IGeneric
    Dim ItemArray() As Object

    Dim d1 As Long, d2 As Long
    Dim ptrKeys As LongPtr, ptrItems As LongPtr

    If (Keys Is Nothing) Or (Items Is Nothing) Then _
        Call System.Throw(Me, "Public Sub SortGenericArrays").IsNothing("(Keys Is Nothing) Or (Items Is Nothing)")
    
    If (Keys.IsEmpty) Or (Items.IsEmpty) Then _
        Call System.Throw(Me, "Public Sub SortGenericArrays").InvalidInput("(Keys.IsEmpty) Or (Items.IsEmpty)")
        
    d1 = Keys.Descriptor
    d2 = Items.Descriptor
    
    ptrKeys = VarPtrArray(KeyArray)
    ptrItems = VarPtrArray(ItemArray)

    Call CopyMemory(Dst:=ByVal ptrKeys, Src:=ByVal VarPtr(d1), Length:=POINTERSIZE): Call CopyMemory(Dst:=ByVal ptrItems, Src:=ByVal VarPtr(d2), Length:=POINTERSIZE)
    Call GenericArray.SortArrayWithItems(Keys:=KeyArray, Items:=ItemArray, LowerBound:=Index, UpperBound:=Length + Index - 1, Comparer:=Comparer, Order:=Order)
    Call ZeroMemory(Dst:=ByVal ptrKeys, Length:=POINTERSIZE): Call ZeroMemory(Dst:=ByVal ptrItems, Length:=POINTERSIZE)

End Sub

Public Sub ShuffleArray(ByRef Element() As IGeneric, ByVal LowerBound As Long, ByVal UpperBound As Long)

    Dim i As Long
    Randomize VBA.Timer
    
    For i = UpperBound To LowerBound Step -1: Call SwapArray(Element, i, Int((UpperBound - LowerBound + 1) * Rnd + LowerBound)): Next

End Sub

Private Sub ShuffleWithItems(ByRef Keys() As IGeneric, ByRef Items() As Object, ByVal LowerBound As Long, ByVal UpperBound As Long)

    Dim i As Long
    Randomize VBA.Timer
    
    For i = UpperBound To LowerBound Step -1: Call SwapWithItem(Keys, Items, i, (i * Rnd + 0.5)): Next

End Sub

Private Sub BinarySort( _
                            ByRef Element() As IGeneric, _
                            ByVal Left As Long, _
                            ByVal Right As Long, _
                            ByVal DepthLimit As Long, _
                            ByVal Order As SortOrder, _
                            ByVal Comparer As IGenericComparer)


    Dim i As Long
    Dim low As Long
    Dim high As Long
    Dim tmp As IGeneric
    


End Sub

Private Sub InsertionSort_Comparer( _
                                        ByRef Elements() As IGeneric, _
                                        ByVal Left As Long, _
                                        ByVal Right As Long, _
                                        ByVal Order As SortOrder, _
                                        ByVal Comparer As IGenericComparer)

    Dim i As Long
    Dim j As Long
    Dim tmp As IGeneric
   
    With Comparer 'Works
        For i = Left To Right - 1
            j = i
            Set tmp = Elements(i + 1)
            Do While (j >= Left)
                If Not .Compare(tmp, Elements(j)) * Order = IsLower Then _
                    Exit Do
             
                Set Elements(j + 1) = Elements(j)
                j = j - 1
            Loop
            Set Elements(j + 1) = tmp
        Next
    End With

End Sub

Private Sub InsertionSort( _
                            ByRef Elements() As IGeneric, _
                            ByVal Left As Long, _
                            ByVal Right As Long, _
                            ByVal Order As SortOrder)

    Dim i As Long
    Dim j As Long
    Dim tmp As IGenericValue 'try cast
    
    If Order = ascending Then
        For i = Left To Right - 1
            j = i
            Set tmp = Elements(i + 1)
            If Not tmp Is Nothing Then
                Do While (j >= Left)
                    If Not tmp.CompareTo(Elements(j)) = IsLower Then _
                        Exit Do
                    
                    Set Elements(j + 1) = Elements(j)
                    j = j - 1
                Loop
            Else
                Do While (j >= Left)
                    Set Elements(j + 1) = Elements(j)
                    j = j - 1
                Loop
            End If
            Set Elements(j + 1) = tmp
        Next
        
    Else
    
        For i = Left To Right - 1
            j = i
            Set tmp = Elements(i + 1)
            If Not tmp Is Nothing Then
                Do While (j >= Left)
                    If Not tmp.CompareTo(Elements(j)) = IsGreater Then _
                        Exit Do
                    
                    Set Elements(j + 1) = Elements(j)
                    j = j - 1
                Loop
            End If
            Set Elements(j + 1) = tmp
        Next
    End If
    
End Sub

Private Sub DepthLimitedQuickSort_Comparer( _
                                            ByRef Element() As IGeneric, _
                                            ByVal Left As Long, _
                                            ByVal Right As Long, _
                                            ByVal DepthLimit As Long, _
                                            ByVal Order As SortOrder, _
                                            ByVal Comparer As IGenericComparer)

    Dim i As Long
    Dim j As Long
    Dim middle As Long
    Dim PartitionSize As Long

    Dim Pivot As IGeneric
    Dim tmp As IGeneric
    
    If (Left > Right) Then _
        Exit Sub
        
    PartitionSize = Right - Left + 1
    
    If (PartitionSize <= INSERTIONSORT_THRESHOLD) Then
    
        If (PartitionSize = 1) Then _
            Exit Sub
        
        If (PartitionSize = 2) Then _
            Call SwapIfGreater_Comparer(Element, Left, Right, Order, Comparer): _
            Exit Sub
        
        If (PartitionSize = 3) Then _
            Call SwapIfGreater_Comparer(Element, Left, Right - 1, Order, Comparer): _
            Call SwapIfGreater_Comparer(Element, Left, Right, Order, Comparer): _
            Call SwapIfGreater_Comparer(Element, Right - 1, Right, Order, Comparer): _
            Exit Sub
        
        Call InsertionSort_Comparer(Elements:=Element, Left:=Left, Right:=Right, Order:=Order, Comparer:=Comparer)
        Exit Sub
    End If
    
    If Order = ascending Then
        Do
            If (DepthLimit = 0) Then _
                Call HeapSort_Comparer(Element, Left, Right, Order, Comparer): Exit Sub
    
            i = Left: j = Right
            middle = i + ((j - i) / 2)
    
            Call SwapIfGreater_Comparer(Element, i, middle, Order, Comparer)
            Call SwapIfGreater_Comparer(Element, i, j, Order, Comparer)
            Call SwapIfGreater_Comparer(Element, middle, j, Order, Comparer)
    
            Set Pivot = Element(middle)
            With Comparer
                Do
                    If (Pivot Is Nothing) Then
                        Do While Not (Element(j) Is Nothing): j = j - 1: Loop
                    Else
                        Do While (.Compare(Pivot, Element(i)) * Order = IsGreater): i = i + 1: Loop
                        Do While (.Compare(Pivot, Element(j)) * Order = IsLower): j = j - 1: Loop
                    End If
                    
                    If (i > j) Then _
                        Exit Do
        
                    If (i < j) Then _
                        Set tmp = Element(i): _
                        Set Element(i) = Element(j): _
                        Set Element(j) = tmp
        
                    i = i + 1: j = j - 1
        
                Loop While (i <= j)
            End With
           
            DepthLimit = DepthLimit - 1
            
            If (j - Left) <= (Right - i) Then
                If (Left < j) Then _
                    Call DepthLimitedQuickSort_Comparer(Element, Left, j, DepthLimit, Order, Comparer)
                Left = i
            Else
                If (i < Right) Then _
                    Call DepthLimitedQuickSort_Comparer(Element, i, Right, DepthLimit, Order, Comparer)
                Right = j
            End If
    
        Loop While (Left < Right)
    
    Else
        
        Do
            If (DepthLimit = 0) Then _
                Call HeapSort_Comparer(Element, Left, Right, Order, Comparer): Exit Sub
    
            i = Left: j = Right
            middle = i + ((j - i) / 2)
    
            Call SwapIfGreater_Comparer(Element, i, middle, Order, Comparer)
            Call SwapIfGreater_Comparer(Element, i, j, Order, Comparer)
            Call SwapIfGreater_Comparer(Element, middle, j, Order, Comparer)
    
            Set Pivot = Element(middle)
            With Comparer
                Do
                    If (Pivot Is Nothing) Then
                        Do While Not (Element(i) Is Nothing): i = i + 1: Loop
                    Else
                        Do While (.Compare(Pivot, Element(i)) = IsLower): i = i + 1: Loop
                        Do While (.Compare(Pivot, Element(j)) = IsGreater): j = j - 1: Loop
                    End If
                    
                    If (i > j) Then _
                        Exit Do
        
                    If (i < j) Then _
                        Set tmp = Element(i): _
                        Set Element(i) = Element(j): _
                        Set Element(j) = tmp
        
                    i = i + 1: j = j - 1
        
                Loop While (i <= j)
            End With
           
            DepthLimit = DepthLimit - 1
            
            If (j - Left) <= (Right - i) Then
                If (Left < j) Then _
                    Call DepthLimitedQuickSort_Comparer(Element, Left, j, DepthLimit, Order, Comparer)
                Left = i
            Else
                If (i < Right) Then _
                    Call DepthLimitedQuickSort_Comparer(Element, i, Right, DepthLimit, Order, Comparer)
                Right = j
            End If
    
        Loop While (Left < Right)
    End If

End Sub

Private Sub HeapSort_Comparer( _
                                ByRef Keys() As IGeneric, _
                                ByVal lo As Long, _
                                ByVal hi As Long, _
                                ByVal Order As SortOrder, _
                                ByVal Comparer As IGenericComparer)

    Dim N As Long: N = hi - lo + 1
    Dim i As Long
    
    If Order = ascending Then
        For i = N / 2 To 1 Step -1
            Call MaxHeap_Comparer(Keys, i, N, lo, Comparer)
        Next
        For i = N To 2 Step -1
            Call SwapArray(Keys, lo, lo + i - 1)
            Call MaxHeap_Comparer(Keys, 1, i - 1, lo, Comparer)
        Next
    Else
        For i = N / 2 To 1 Step -1
            Call MinHeap_Comparer(Keys, i, N, lo, Comparer)
        Next
        For i = N To 2 Step -1
            Call SwapArray(Keys, lo, lo + i - 1)
            Call MinHeap_Comparer(Keys, 1, i - 1, lo, Comparer)
        Next
    End If
    
End Sub

Private Sub MaxHeap_Comparer(ByRef Keys() As IGeneric, ByVal i As Long, ByVal N As Long, ByVal lo As Long, ByVal Comparer As IGenericComparer)

    Dim d As IGeneric: Set d = Keys(lo + i - 1)
    Dim Child As Long
    
    With Comparer
        Do While (i <= N / 2)
            Child = 2 * i
            If (Child < N) Then
                If (Keys(lo + Child - 1) Is Nothing) Then
                    Child = Child + 1
                ElseIf (.Compare(Keys(lo + Child - 1), Keys(lo + Child)) = IsLower) Then
                    Child = Child + 1
                End If
            End If
            
            If (Keys(lo + Child - 1) Is Nothing) Then _
                Exit Do
    
            If (.Compare(Keys(lo + Child - 1), d) = IsLower) Then _
                Exit Do
    
            Set Keys(lo + i - 1) = Keys(lo + Child - 1)
            i = Child
        Loop
    End With
    
    Set Keys(lo + i - 1) = d

End Sub

Private Sub MinHeap_Comparer(ByRef Keys() As IGeneric, ByVal i As Long, ByVal N As Long, ByVal lo As Long, ByVal Comparer As IGenericComparer)

    Dim d As IGeneric: Set d = Keys(lo + i - 1)
    Dim Child As Long
    
    With Comparer
        Do While (i <= N / 2)
            Child = 2 * i
            If (Child < N) Then _
                If (.Compare(Keys(lo + Child - 1), Keys(lo + Child)) = IsGreater) Then _
                    Child = Child + 1
    
            If Not (Keys(lo + Child - 1) Is Nothing) Then _
                If (.Compare(Keys(lo + Child - 1), d) = IsGreater) Then _
                    Exit Do
            
            Set Keys(lo + i - 1) = Keys(lo + Child - 1)
            i = Child
        Loop
    End With
    
    Set Keys(lo + i - 1) = d

End Sub

Private Sub DepthLimitedQuickSort( _
                                    ByRef Element() As IGeneric, _
                                    ByVal Left As Long, _
                                    ByVal Right As Long, _
                                    ByVal DepthLimit As Long, _
                                    ByVal Order As SortOrder)

    Dim i As Long
    Dim j As Long
    Dim middle As Long
    Dim PartitionSize As Long
    
    Dim Pivot As IGenericValue
    Dim tmp As IGeneric
    
    If (Left > Right) Then _
        Exit Sub
        
    PartitionSize = Right - Left + 1
    
    If (PartitionSize <= INSERTIONSORT_THRESHOLD) Then

        If (PartitionSize = 1) Then _
            Exit Sub

        If (PartitionSize = 2) Then _
            Call SwapIfGreater(Element, Left, Right, Order): _
            Exit Sub

        If (PartitionSize = 3) Then _
            Call SwapIfGreater(Element, Left, Right - 1, Order): _
            Call SwapIfGreater(Element, Left, Right, Order): _
            Call SwapIfGreater(Element, Right - 1, Right, Order): _
            Exit Sub

        Call InsertionSort(Elements:=Element, Left:=Left, Right:=Right, Order:=Order)
        Exit Sub
    End If

    If Order = ascending Then
        Do
            If (DepthLimit = 0) Then _
                Call HeapSort(Element, Left, Right, Order): Exit Sub
    
            i = Left: j = Right
            middle = i + ((j - i) / 2)
    
            Call SwapIfGreater(Element, i, middle, Order)
            Call SwapIfGreater(Element, i, j, Order)
            Call SwapIfGreater(Element, middle, j, Order)
    
            Set Pivot = Element(middle) 'try to cast
            With Pivot
                Do
            
                    If (Pivot Is Nothing) Then
                        Do While Not (Element(j) Is Nothing): j = j - 1: Loop
                    Else
                        Do While (.CompareTo(Element(i)) = IsGreater): i = i + 1: Loop
                        Do While (.CompareTo(Element(j)) = IsLower): j = j - 1: Loop
                    End If
        
                    If (i > j) Then _
                        Exit Do
        
                    If (i < j) Then _
                        Set tmp = Element(i): _
                        Set Element(i) = Element(j): _
                        Set Element(j) = tmp
        
                    i = i + 1: j = j - 1
        
                Loop While (i <= j)
            End With
            
            DepthLimit = DepthLimit - 1
    
            If (j - Left) <= (Right - i) Then
                If (Left < j) Then _
                    Call DepthLimitedQuickSort(Element, Left, j, DepthLimit, Order)
                Left = i
            Else
                If (i < Right) Then _
                    Call DepthLimitedQuickSort(Element, i, Right, DepthLimit, Order)
                Right = j
            End If
    
        Loop While (Left < Right)
        
    Else
    
        Do
            If (DepthLimit = 0) Then _
                Call HeapSort(Element, Left, Right, Order): Exit Sub
    
            i = Left: j = Right
            middle = i + ((j - i) / 2)
    
            Call SwapIfGreater(Element, i, middle, Order)
            Call SwapIfGreater(Element, i, j, Order)
            Call SwapIfGreater(Element, middle, j, Order)
    
            Set Pivot = Element(middle) 'try to cast
            With Pivot
                Do
            
                    If (Pivot Is Nothing) Then
                        Do While Not (Element(i) Is Nothing): i = i + 1: Loop
                    Else
                        Do While (.CompareTo(Element(i)) = IsLower): i = i + 1: Loop
                        Do While (.CompareTo(Element(j)) = IsGreater): j = j - 1: Loop
                    End If
        
                    If (i > j) Then _
                        Exit Do
        
                    If (i < j) Then _
                        Set tmp = Element(i): _
                        Set Element(i) = Element(j): _
                        Set Element(j) = tmp
        
                    i = i + 1: j = j - 1
        
                Loop While (i <= j)
            End With
            
            DepthLimit = DepthLimit - 1
    
            If (j - Left) <= (Right - i) Then
                If (Left < j) Then _
                    Call DepthLimitedQuickSort(Element, Left, j, DepthLimit, Order)
                Left = i
            Else
                If (i < Right) Then _
                    Call DepthLimitedQuickSort(Element, i, Right, DepthLimit, Order)
                Right = j
            End If
    
        Loop While (Left < Right)
    End If
   
End Sub
                              
Private Sub HeapSort( _
                            ByRef Keys() As IGeneric, _
                            ByVal lo As Long, _
                            ByVal hi As Long, _
                            ByVal Order As SortOrder)

    Dim N As Long: N = hi - lo + 1
    Dim i As Long
    
    If Order = ascending Then
        For i = N / 2 To 1 Step -1
            Call MaxHeap(Keys, i, N, lo)
        Next
        For i = N To 2 Step -1
            Call SwapArray(Keys, lo, lo + i - 1)
            Call MaxHeap(Keys, 1, i - 1, lo)
        Next
    Else
        For i = N / 2 To 1 Step -1
            Call MinHeap(Keys, i, N, lo)
        Next
        For i = N To 2 Step -1
            Call SwapArray(Keys, lo, lo + i - 1)
            Call MinHeap(Keys, 1, i - 1, lo)
        Next
    End If
End Sub

Private Sub MaxHeap(ByRef Keys() As IGeneric, ByVal i As Long, ByVal N As Long, ByVal lo As Long)

    Dim d As IGenericValue: Set d = Keys(lo + i - 1) 'try to cast
    Dim Element As IGenericValue
    Dim Child As Long

    Do While (i <= N / 2)
        Child = 2 * i
        Set Element = Keys(lo + Child - 1)
        If (Child < N) Then
            If Element Is Nothing Then
                Child = Child + 1: Set Element = Keys(lo + Child - 1)
            ElseIf (Element.CompareTo(Keys(lo + Child)) = IsLower) Then
                Child = Child + 1: Set Element = Keys(lo + Child - 1)
            End If
        End If

        If (Element Is Nothing) Then _
            Exit Do

        If (Element.CompareTo(d) = IsLower) Then _
            Exit Do

        Set Keys(lo + i - 1) = Element
        i = Child
    Loop

    Set Keys(lo + i - 1) = d

End Sub

Private Sub MinHeap(ByRef Keys() As IGeneric, ByVal i As Long, ByVal N As Long, ByVal lo As Long)

    Dim d As IGenericValue: Set d = Keys(lo + i - 1) 'try to cast
    Dim Element As IGenericValue
    Dim Child As Long

    Do While (i <= N / 2)
        Child = 2 * i
        Set Element = Keys(lo + Child - 1)
        If (Child < N) Then _
            If Not (Element Is Nothing) Then _
                If (Element.CompareTo(Keys(lo + Child)) = IsGreater) Then _
                    Child = Child + 1: Set Element = Keys(lo + Child - 1)
    
        If Not (Element Is Nothing) Then _
            If (Element.CompareTo(d) = IsGreater) Then _
                Exit Do

        Set Keys(lo + i - 1) = Element
        i = Child
    Loop

    Set Keys(lo + i - 1) = d

End Sub
                      
Private Sub DepthLimitedQuickSortWithItems( _
                                            ByRef Keys() As IGeneric, _
                                            ByRef Items() As Object, _
                                            ByVal Left As Long, _
                                            ByVal Right As Long, _
                                            ByVal DepthLimit As Long, _
                                            ByVal Order As SortOrder)

    Dim i As Long
    Dim j As Long
    Dim middle As Long
    Dim PartitionSize As Long

    Dim Pivot As IGenericValue
    Dim tmp As IGeneric

    If (Left > Right) Then _
        Exit Sub
        
    PartitionSize = Right - Left + 1
    
    If (PartitionSize <= INSERTIONSORT_THRESHOLD) Then
        
        If (PartitionSize = 1) Then _
            Exit Sub
        
        If (PartitionSize = 2) Then _
            Call SwapIfGreaterWithItem(Keys, Items, Left, Right, Order): _
            Exit Sub
        
        If (PartitionSize = 3) Then _
            Call SwapIfGreaterWithItem(Keys, Items, Left, Right - 1, Order): _
            Call SwapIfGreaterWithItem(Keys, Items, Left, Right, Order): _
            Call SwapIfGreaterWithItem(Keys, Items, Right - 1, Right, Order): _
            Exit Sub
            
        Call InsertionSort(Elements:=Keys, Left:=Left, Right:=Right, Order:=Order)
        Exit Sub
        
    End If
    
    If Order = ascending Then
        Do
            If (DepthLimit = 0) Then _
                Call HeapSortWithItems(Keys, Items, Left, Right, Order): Exit Sub
    
            i = Left: j = Right
            middle = i + ((j - i) / 2)
    
            Call SwapIfGreaterWithItem(Keys, Items, i, middle, Order)
            Call SwapIfGreaterWithItem(Keys, Items, i, j, Order)
            Call SwapIfGreaterWithItem(Keys, Items, middle, j, Order)
    
            Set Pivot = Keys(middle) 'try to cast
            With Pivot
                Do
                
                    If (Pivot Is Nothing) Then
                        Do While Not (Keys(j) Is Nothing): j = j - 1: Loop
                    Else
                        Do While (.CompareTo(Keys(i)) = IsGreater): i = i + 1: Loop
                        Do While (.CompareTo(Keys(j)) = IsLower): j = j - 1: Loop
                    End If
        
                    If (i > j) Then _
                        Exit Do
                
                    If (i < j) Then _
                        Set tmp = Keys(i): Set Keys(i) = Keys(j): Set Keys(j) = tmp: _
                        Set tmp = Items(i): Set Items(i) = Items(j): Set Items(j) = tmp
                
                    i = i + 1: j = j - 1
                    
                Loop While (i <= j)
            End With
            
            DepthLimit = DepthLimit - 1
    
            If (j - Left) <= (Right - i) Then
                If (Left < j) Then _
                    Call DepthLimitedQuickSortWithItems(Keys, Items, Left, j, DepthLimit, Order)
                Left = i
            Else
                If (i < Right) Then _
                    Call DepthLimitedQuickSortWithItems(Keys, Items, i, Right, DepthLimit, Order)
                Right = j
            End If
    
        Loop While (Left < Right)
        
    Else
    
        Do
            If (DepthLimit = 0) Then _
                Call HeapSortWithItems(Keys, Items, Left, Right, Order): Exit Sub
    
            i = Left: j = Right
            middle = i + ((j - i) / 2)
    
            Call SwapIfGreaterWithItem(Keys, Items, i, middle, Order)
            Call SwapIfGreaterWithItem(Keys, Items, i, j, Order)
            Call SwapIfGreaterWithItem(Keys, Items, middle, j, Order)
    
            Set Pivot = Keys(middle) 'try to cast
            With Pivot
                Do
                
                    If (Pivot Is Nothing) Then
                        Do While Not (Keys(i) Is Nothing): i = i + 1: Loop
                    Else
                        Do While (.CompareTo(Keys(i)) = IsLower): i = i + 1: Loop
                        Do While (.CompareTo(Keys(j)) = IsGreater): j = j - 1: Loop
                    End If
        
                    If (i > j) Then _
                        Exit Do
                
                    If (i < j) Then _
                        Set tmp = Keys(i): Set Keys(i) = Keys(j): Set Keys(j) = tmp: _
                        Set tmp = Items(i): Set Items(i) = Items(j): Set Items(j) = tmp
                
                    i = i + 1: j = j - 1
                    
                Loop While (i <= j)
            End With
            
            DepthLimit = DepthLimit - 1
    
            If (j - Left) <= (Right - i) Then
                If (Left < j) Then _
                    Call DepthLimitedQuickSortWithItems(Keys, Items, Left, j, DepthLimit, Order)
                Left = i
            Else
                If (i < Right) Then _
                    Call DepthLimitedQuickSortWithItems(Keys, Items, i, Right, DepthLimit, Order)
                Right = j
            End If
    
        Loop While (Left < Right)
    End If
        
End Sub

Private Sub HeapSortWithItems( _
                                ByRef Keys() As IGeneric, _
                                ByRef Items() As Object, _
                                ByVal lo As Long, _
                                ByVal hi As Long, _
                                ByVal Order As SortOrder)

    Dim N As Long: N = hi - lo + 1
    Dim i As Long
    
    If Order = ascending Then
        For i = N / 2 To 1 Step -1
            Call MaxHeapWithItems(Keys, Items, i, N, lo)
        Next
        For i = N To 2 Step -1
            Call SwapWithItem(Keys, Items, lo, lo + i - 1)
            Call MaxHeapWithItems(Keys, Items, 1, i - 1, lo)
        Next
    Else
        For i = N / 2 To 1 Step -1
            Call MinHeapWithItems(Keys, Items, i, N, lo)
        Next
        For i = N To 2 Step -1
            Call SwapWithItem(Keys, Items, lo, lo + i - 1)
            Call MinHeapWithItems(Keys, Items, 1, i - 1, lo)
        Next
    End If
    
End Sub

Private Sub MaxHeapWithItems( _
                                ByRef Keys() As IGeneric, _
                                ByRef Items() As Object, _
                                ByVal i As Long, _
                                ByVal N As Long, _
                                ByVal lo As Long)

    Dim d As IGenericValue: Set d = Keys(lo + i - 1)
    Dim dItem As Object: Set dItem = Items(lo + i - 1)
    Dim Element As IGenericValue
    Dim Child As Long

    Do While (i <= N / 2)
        Child = 2 * i
        Set Element = Keys(lo + Child - 1)
        If (Child < N) Then
            If (Element Is Nothing) Then
                Child = Child + 1: Set Element = Keys(lo + Child - 1)
            ElseIf (Element.CompareTo(Keys(lo + Child)) = IsLower) Then
                Child = Child + 1: Set Element = Keys(lo + Child - 1)
            End If
        End If

        If (Element Is Nothing) Then _
            Exit Do

        If (Element.CompareTo(d) = IsLower) Then _
            Exit Do

        Set Keys(lo + i - 1) = Element
        Set Items(lo + i - 1) = Items(lo + Child - 1)
        i = Child
    Loop

    Set Keys(lo + i - 1) = d
    Set Items(lo + i - 1) = dItem

End Sub

Private Sub MinHeapWithItems( _
                                ByRef Keys() As IGeneric, _
                                ByRef Items() As Object, _
                                ByVal i As Long, _
                                ByVal N As Long, _
                                ByVal lo As Long)

    Dim d As IGenericValue: Set d = Keys(lo + i - 1)
    Dim dItem As Object: Set dItem = Items(lo + i - 1)
    Dim Element As IGenericValue
    Dim Child As Long

    Do While (i <= N / 2)
        Child = 2 * i
        Set Element = Keys(lo + Child - 1)
        If (Child < N) Then _
            If Not (Element Is Nothing) Then _
                If (Element.CompareTo(Keys(lo + Child)) = IsGreater) Then _
                    Child = Child + 1: Set Element = Keys(lo + Child - 1)
    
        If Not (Element Is Nothing) Then _
            If (Element.CompareTo(d) = IsGreater) Then _
                Exit Do
        
        Set Keys(lo + i - 1) = Element
        Set Items(lo + i - 1) = Items(lo + Child - 1)
        
        i = Child
    Loop

    Set Keys(lo + i - 1) = d
    Set Items(lo + i - 1) = dItem
    
End Sub

Private Sub DepthLimitedQuickSortWithItems_Comparer( _
                                                        ByRef Keys() As IGeneric, _
                                                        ByRef Items() As Object, _
                                                        ByVal Left As Long, _
                                                        ByVal Right As Long, _
                                                        ByVal DepthLimit As Long, _
                                                        ByVal Order As SortOrder, _
                                                        ByVal Comparer As IGenericComparer)

    Dim i As Long
    Dim j As Long
    Dim middle As Long
    Dim PartitionSize As Long

    Dim Pivot As IGeneric
    Dim tmp As IGeneric

    If (Left > Right) Then _
        Exit Sub
        
    PartitionSize = Right - Left + 1
    
    If (PartitionSize <= INSERTIONSORT_THRESHOLD) Then
         
        If (PartitionSize = 1) Then _
            Exit Sub
        
        If (PartitionSize = 2) Then _
            Call SwapIfGreaterWithItem_Comparer(Keys, Items, Left, Right, Order, Comparer): _
            Exit Sub
        
        If (PartitionSize = 3) Then _
            Call SwapIfGreaterWithItem_Comparer(Keys, Items, Left, Right - 1, Order, Comparer): _
            Call SwapIfGreaterWithItem_Comparer(Keys, Items, Left, Right, Order, Comparer): _
            Call SwapIfGreaterWithItem_Comparer(Keys, Items, Right - 1, Right, Order, Comparer): _
            Exit Sub
            
        Call InsertionSort_Comparer(Elements:=Keys, Left:=Left, Right:=Right, Order:=Order, Comparer:=Comparer)
        Exit Sub
        
    End If
    
    If Order = ascending Then
        Do
            If (DepthLimit = 0) Then _
                Call HeapSortWithItems_Comparer(Keys, Items, Left, Right, Order, Comparer): Exit Sub
    
            i = Left: j = Right
            middle = i + ((j - i) / 2)
    
            Call SwapIfGreaterWithItem_Comparer(Keys, Items, i, middle, Order, Comparer)
            Call SwapIfGreaterWithItem_Comparer(Keys, Items, i, j, Order, Comparer)
            Call SwapIfGreaterWithItem_Comparer(Keys, Items, middle, j, Order, Comparer)
    
            Set Pivot = Keys(middle)
            With Comparer
                Do
                    
                    If (Pivot Is Nothing) Then
                        Do While Not (Keys(j) Is Nothing): j = j - 1: Loop
                    Else
                        Do While (.Compare(Pivot, Keys(i)) * Order = IsGreater): i = i + 1: Loop
                        Do While (.Compare(Pivot, Keys(j)) * Order = IsLower): j = j - 1: Loop
                    End If
                    
                    If (i > j) Then _
                        Exit Do
                
                    If (i < j) Then _
                        Set tmp = Keys(i): Set Keys(i) = Keys(j): Set Keys(j) = tmp: _
                        Set tmp = Items(i): Set Items(i) = Items(j): Set Items(j) = tmp
                
                    i = i + 1: j = j - 1
                    
                Loop While (i <= j)
            End With
    
            DepthLimit = DepthLimit - 1
    
            If (j - Left) <= (Right - i) Then
                If (Left < j) Then _
                    Call DepthLimitedQuickSortWithItems_Comparer(Keys, Items, Left, j, DepthLimit, Order, Comparer)
                Left = i
            Else
                If (i < Right) Then _
                    Call DepthLimitedQuickSortWithItems_Comparer(Keys, Items, i, Right, DepthLimit, Order, Comparer)
                Right = j
            End If
    
        Loop While (Left < Right)
    
    Else
    
        Do
            If (DepthLimit = 0) Then _
                Call HeapSortWithItems_Comparer(Keys, Items, Left, Right, Order, Comparer): Exit Sub
    
            i = Left: j = Right
            middle = i + ((j - i) / 2)
    
            Call SwapIfGreaterWithItem_Comparer(Keys, Items, i, middle, Order, Comparer)
            Call SwapIfGreaterWithItem_Comparer(Keys, Items, i, j, Order, Comparer)
            Call SwapIfGreaterWithItem_Comparer(Keys, Items, middle, j, Order, Comparer)
    
            Set Pivot = Keys(middle)
            With Comparer
                Do
                    If (Pivot Is Nothing) Then
                        Do While Not (Keys(i) Is Nothing): i = i + 1: Loop
                    Else
                        Do While (.Compare(Pivot, Keys(i)) = IsLower): i = i + 1: Loop
                        Do While (.Compare(Pivot, Keys(j)) = IsGreater): j = j - 1: Loop
                    End If
                    
                    If (i > j) Then _
                        Exit Do
                
                    If (i < j) Then _
                        Set tmp = Keys(i): Set Keys(i) = Keys(j): Set Keys(j) = tmp: _
                        Set tmp = Items(i): Set Items(i) = Items(j): Set Items(j) = tmp
                
                    i = i + 1: j = j - 1
                    
                Loop While (i <= j)
            End With
    
            DepthLimit = DepthLimit - 1
    
            If (j - Left) <= (Right - i) Then
                If (Left < j) Then _
                    Call DepthLimitedQuickSortWithItems_Comparer(Keys, Items, Left, j, DepthLimit, Order, Comparer)
                Left = i
            Else
                If (i < Right) Then _
                    Call DepthLimitedQuickSortWithItems_Comparer(Keys, Items, i, Right, DepthLimit, Order, Comparer)
                Right = j
            End If
    
        Loop While (Left < Right)
    End If
End Sub

Private Sub HeapSortWithItems_Comparer( _
                                        ByRef Keys() As IGeneric, _
                                        ByRef Items() As Object, _
                                        ByVal lo As Long, _
                                        ByVal hi As Long, _
                                        ByVal Order As SortOrder, _
                                        ByVal Comparer As IGenericComparer)

    Dim N As Long: N = hi - lo + 1
    Dim i As Long
    
    If Order = ascending Then
        For i = N / 2 To 1 Step -1
            Call MaxHeapWithItems_Comparer(Keys, Items, i, N, lo, Comparer)
        Next
        For i = N To 2 Step -1
            Call SwapWithItem(Keys, Items, lo, lo + i - 1)
            Call MaxHeapWithItems_Comparer(Keys, Items, 1, i - 1, lo, Comparer)
        Next
    Else
        For i = N / 2 To 1 Step -1
            Call MinHeapWithItems_Comparer(Keys, Items, i, N, lo, Comparer)
        Next
        For i = N To 2 Step -1
            Call SwapWithItem(Keys, Items, lo, lo + i - 1)
            Call MinHeapWithItems_Comparer(Keys, Items, 1, i - 1, lo, Comparer)
        Next
    End If
    
End Sub

Private Sub MaxHeapWithItems_Comparer( _
                                        ByRef Keys() As IGeneric, _
                                        ByRef Items() As Object, _
                                        ByVal i As Long, _
                                        ByVal N As Long, _
                                        ByVal lo As Long, _
                                        ByVal Comparer As IGenericComparer)

    Dim d As IGeneric: Set d = Keys(lo + i - 1)
    Dim dItem As Object: Set dItem = Items(lo + i - 1)
    Dim Child As Long
    
    With Comparer
        Do While (i <= N / 2)
            Child = 2 * i
            If (Child < N) Then
                If (Keys(lo + Child - 1) Is Nothing) Then
                    Child = Child + 1
                ElseIf (.Compare(Keys(lo + Child - 1), Keys(lo + Child)) = IsLower) Then
                    Child = Child + 1
                End If
            End If
    
            If (Keys(lo + Child - 1) Is Nothing) Then _
                Exit Do
    
            If (.Compare(Keys(lo + Child - 1), d) = IsLower) Then _
                Exit Do
    
            Set Keys(lo + i - 1) = Keys(lo + Child - 1)
            Set Items(lo + i - 1) = Items(lo + Child - 1)
            i = Child
        Loop
    End With
    
    Set Keys(lo + i - 1) = d
    Set Items(lo + i - 1) = dItem

End Sub

Private Sub MinHeapWithItems_Comparer( _
                                        ByRef Keys() As IGeneric, _
                                        ByRef Items() As Object, _
                                        ByVal i As Long, _
                                        ByVal N As Long, _
                                        ByVal lo As Long, _
                                        ByVal Comparer As IGenericComparer)

    Dim d As IGeneric: Set d = Keys(lo + i - 1)
    Dim dItem As Object: Set dItem = Items(lo + i - 1)
    Dim Child As Long

    With Comparer
        Do While (i <= N / 2)
            Child = 2 * i
            If (Child < N) Then
                If Not (Keys(lo + Child - 1) Is Nothing) Then
                    If (.Compare(Keys(lo + Child - 1), Keys(lo + Child)) = IsGreater) Then _
                        Child = Child + 1
        
                End If
            End If
    
            If Not (Keys(lo + Child - 1) Is Nothing) Then _
                If (.Compare(Keys(lo + Child - 1), d) = IsGreater) Then _
                    Exit Do
            
            Set Keys(lo + i - 1) = Keys(lo + Child - 1)
            Set Items(lo + i - 1) = Items(lo + Child - 1)
            i = Child
        Loop
    End With
    
    Set Keys(lo + i - 1) = d
    Set Items(lo + i - 1) = dItem
    
End Sub
Private Sub Swap(ByVal i As Long, j As Long): Call SwapArray(myclass.Internal, i, j): End Sub
Private Sub SwapArray(ByRef A() As IGeneric, ByVal i As Long, ByVal j As Long)

    Dim t As IGeneric
    
    Set t = A(i): Set A(i) = A(j): Set A(j) = t

End Sub

Private Sub SwapIfGreater(ByRef Keys() As IGeneric, ByVal A As Long, ByVal B As Long, ByVal Order As SortOrder)
    
    Dim t As IGeneric
    Dim Element As IGenericValue
    
    If Not (A = B) Then
        Set Element = Keys(A)
        If Not (Element Is Nothing) Then
            If (Element.CompareTo(Keys(B)) * Order = IsGreater) Then
                Set t = Keys(A)
                Set Keys(A) = Keys(B)
                Set Keys(B) = t
            End If
        End If
    End If
    
End Sub

Private Sub SwapIfGreaterWithItem(ByRef Keys() As IGeneric, ByRef Items() As Object, ByVal A As Long, ByVal B As Long, ByVal Order As SortOrder)
    
    Dim t As IGeneric
    Dim Element As IGenericValue
   
    If Not (A = B) Then
        Set Element = Keys(A)
        If Not (Element Is Nothing) Then
            If (Element.CompareTo(Keys(B)) * Order = IsGreater) Then
                Set t = Keys(A): Set Keys(A) = Keys(B): Set Keys(B) = t
                Set t = Items(A): Set Items(A) = Items(B): Set Items(B) = t
            End If
        End If
    End If
    
End Sub

Private Sub SwapIfGreaterWithItem_Comparer(ByRef Keys() As IGeneric, ByRef Items() As Object, ByVal A As Long, ByVal B As Long, ByVal Order As SortOrder, ByVal Comparer As IGenericComparer)
    
    Dim t As IGeneric
   
    If Not (A = B) Then
        If Not (Keys(A) Is Nothing) Then
            If (Comparer.Compare(Keys(A), Keys(B)) * Order = IsGreater) Then
                Set t = Keys(A): Set Keys(A) = Keys(B): Set Keys(B) = t
                Set t = Items(A): Set Items(A) = Items(B): Set Items(B) = t
            End If
        End If
    End If
    
End Sub

Private Sub SwapIfGreater_Comparer(ByRef Keys() As IGeneric, ByVal A As Long, ByVal B As Long, ByVal Order As SortOrder, ByVal Comparer As IGenericComparer)
    
    Dim t As IGeneric
   
    If Not (A = B) Then
        If Not (Keys(A) Is Nothing) Then
            If (Comparer.Compare(Keys(A), Keys(B)) * Order = IsGreater) Then
                Set t = Keys(A)
                Set Keys(A) = Keys(B)
                Set Keys(B) = t
            End If
        End If
    End If
    
End Sub

Private Sub SwapWithItem(ByRef A() As IGeneric, ByRef B() As Object, ByVal i As Long, ByVal j As Long)
    
    Dim t As IGeneric
    
    Set t = A(i): Set A(i) = A(j): Set A(j) = t
    Set t = B(i): Set B(i) = B(j): Set B(j) = t
    
End Sub

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericArray: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GenericArray: End Property

Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericArray):  End Property

Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = Me.Elements.Copy: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    If Other Is Nothing Then _
        IGeneric_Equals = False: Exit Function
    
    IGeneric_Equals = (Me Is Other)
    
End Function

Private Function IGeneric_ToString() As String
    IGeneric_ToString = _
                        "Class: " & TypeName$(Me) & vbNewLine & _
                        "Rank: " & Me.Rank & vbNewLine & _
                        "Elements: " & Me.Length & vbNewLine & _
                        "MemorySize: " & Me.Size & " Byte"
End Function

'IGenericReadOnlyList
Private Property Get IGenericReadOnlyList_Elements() As IGenericCollection: Set IGenericReadOnlyList_Elements = Me: End Property
Private Property Get IGenericReadOnlyList_First() As Long: IGenericReadOnlyList_First = MIN_INDEX: End Property
Private Property Get IGenericReadOnlyList_Last() As Long: IGenericReadOnlyList_Last = myclass.Elements - 1: End Property
Private Property Get IGenericReadOnlyList_ElementAt(ByVal Index As Long) As IGeneric
    
    If (Index >= myclass.Elements) Then _
        Call System.Throw(Me, "Private Property Get IGenericReadOnlyList_ElementAt").IndexOutofRange("(Index >= myclass.Elements)")

    Set IGenericReadOnlyList_ElementAt = myclass.Internal(Index)
        
End Property
Private Function IGenericReadOnlyList_IndexOf(ByVal Element As IGeneric) As Long: IGenericReadOnlyList_IndexOf = Me.IndexOf(Value:=Element, Index:=MIN_INDEX, Length:=myclass.Elements): End Function

'IGenericCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = myclass.Elements: End Property
Private Sub IGenericCollection_Clear()
    
    Dim i As Long
    
    If Me.IsEmpty Then _
        Exit Sub
    
    myclass.Struct.pvData = 0
    Call Catch(SafeArrayDestroyData(ByVal myclass.Descriptor)) 'if not fFeatures & FADF_STATIC then pvData-> null
    Call Catch(SafeArrayAllocData(ByVal myclass.Descriptor))
    
    For i = UBound(myclass.Indices) To LBound(myclass.Indices) Step -1
        myclass.Indices(i) = MIN_INDEX
    Next i
    
    Call Catch(SafeArrayPtrOfIndex(ByVal myclass.Descriptor, myclass.Indices(i + 1), myclass.Struct.pvData))
    
End Sub
'TODO Rewrite CopyArrays -> CopyTo
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call Me.CopyTo(GArray:=GArray, ArrayIndex:=ArrayIndex, Index:=MIN_INDEX, Length:=myclass.Elements): End Sub

Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = Me.Elements.Copy: End Function

Private Function IGenericCollection_Contains(ByVal Item As IGeneric) As Boolean
    
    Dim Element As IGeneric
    If Me.Rank = 1 Then
        IGenericCollection_Contains = Me.IndexOf(Value:=Item, Index:=MIN_INDEX, Length:=myclass.Elements) >= MIN_INDEX:
    Else
        With Me.Elements.Iterator
            Do While .HasNext(Element)
                If (Item.Equals(Element)) Then _
                    IGenericCollection_Contains = True: Exit Function
            Loop
        End With
    End If

End Function

Private Function IGenericCollection_Copy() As IGenericCollection
    
    Dim Clone As Member
    Dim i As Long
    
    If (Me.IsEmpty) Then _
        Call System.Throw(Me, "Public Function Copy").InvalidOperation("Me.IsEmpty")
  
    Clone.Elements = myclass.Elements
    Clone.Struct.cDims = myclass.Struct.cDims
    Clone.Struct.fFeatures = myclass.Struct.fFeatures
    Clone.Struct.cbElements = myclass.Struct.cbElements
    Clone.Struct.Bounds = myclass.Struct.Bounds
    Clone.Indices = myclass.Indices

    Call Catch(SafeArrayCopy(ByVal myclass.Descriptor, Clone.Descriptor))
    
    If (Clone.Descriptor = 0) Then _
        Call System.Throw(Me, "Public Function Copy").InvalidOperation("Descriptor = 0")
    Call LoadArray(Clone.Internal, Clone.Descriptor)
    
    For i = LBound(Clone.Indices) To UBound(Clone.Indices)
        Clone.Indices(i) = MIN_INDEX
    Next
    
    Call Catch(SafeArrayPtrOfIndex(ByVal Clone.Descriptor, Clone.Indices(0), Clone.Struct.pvData))

    Set IGenericCollection_Copy = System.CreateInstance(New GenericArray, VarPtr(Clone), LenB(Clone))
    
End Function

Private Function IGenericCollection_Iterator() As IGenericIterator

    Dim c As Member

    Call LoadIterator(c.Iterator, Me)
    Set IGenericCollection_Iterator = System.CreateInstance(New GenericArray, VarPtr(c), LenB(c))

End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Iterator.current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    With myclass.Iterator
        If (.Rank = 1) Then
            If .Index < .EndIndex Then
                Set .current = .Matrix.ElementAt(.Index)
                .Index = .Index + 1
                Set Out = .current
                IGenericIterator_HasNext = True
                Exit Function
            End If
           
        Else
   
            Call System.Throw(Me, "Private Function IGenericIterator_HasNext").NotImplemented("multi dim not implemented yet")
            
        End If
    
        If .HasLoaded = False Then _
            Call System.Throw(Me, "Private Function IGenericIterator_HasNext").InvalidIterator
    
        Set .current = Nothing
        Set Out = .current
        
    End With
End Function

Private Sub IGenericIterator_Reset()
    With myclass.Iterator
        If .HasLoaded = False Then _
            Call System.Throw(Me, "Private Sub IGenericIterator_Reset").InvalidIterator
            
        Call LoadIterator(myclass.Iterator, .Matrix)
    End With
End Sub





