VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "Collections"
'@IgnoreModule ProcedureCanBeWrittenAsFunction
 Option Explicit

'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  GenericArray
'*
'* Purpose: An array stores a fixed-size sequential collection of elements of the IGeneric type which are accessed through computed indices.
'*          The variables contained in an array, also called the elements of the array.
'*          An array has a rank which determines the number of indices associated with each array element.
'*          The rank of an array is also referred to as the dimensions of the array.
'*          An array with a rank of one is called a single-dimensional array. An array with a rank greater than one is called a multi-dimensional array.
'*          Specific sized multi-dimensional arrays are often referred to as two-dimensional arrays, three-dimensional arrays, and so on.
'*          Each dimension of an array has an associated length which is an integral number greater than or equal to zero.
'*          The dimension lengths are not part of the type of the array, but rather are established when an instance of the array is created at run-time.
'*          The length of a dimension determines the valid range of indices for that dimension:
'*              -> For a dimension of length N, indices can range from 0 to N - 1 inclusive.
'*          The total number of elements in an array is the product of the lengths of each dimension in the array.
'*          If one or more of the dimensions of an array have a length of zero, the array is said to be empty.
'*          It provides methods for creating, manipulating, searching, and sorting arrays, thereby serving as the base class for all Array instances.
'*          An element is a (reference) value in an Array. The length of an Array is the total number of elements it can contain.
'*          The lower bound of an Array is the index of its first element (zero by default).
'*          A multidimensional Array can have different upper bounds for each dimension.
'*
'*          The Array class is part of the IGenericCollection namespaces. Unlike the other classes which are the IGenericCollection implemented,
'*          Array has a fixed capacity. To increase the capacity, you must create a new Array object with the required capacity,
'*          copy the elements from the old Array object to the new one, and delete the old Array.
'*          The Array is not guaranteed to be sorted
'*              -> You must sort the Array prior to performing operations (such as BinarySearch, Distinct) that require the Array to be sorted.
'*
'* Notes:   An array can be Single-Dimensional, Multidimensional or Jagged and can have a maximum of 32 dimensions.
'*          The number of dimensions and the length of each dimension are established when the array instance is created
'*              -> These values can't be changed during the lifetime of the instance!
'*          The default values (reference elements) of an array elements are set to null.
'*          Each element can be referred to by an index.
'*          The length of the array is the total number of item present in the array.
'*          Arrays are zero indexed: an array with n elements is indexed from 0 to n-1.
'*          Array elements can be of any IGeneric type, including an GenericArray -> A jagged array is an array of arrays, and therefore its elements are reference types and are initialized to null.
'*
'*          If you have an array of types (such as string or integer) that already support IGenericValue, In that case, the elements of the array are cast to the default implementation of IGenericValue
'*          You can sort that array without providing any explicit reference to an IGenericComparer!
'*
'*
'*
'*
'*======================================================================================================================

Implements IGeneric
Implements IGenericCollection
Implements IGenericSequence

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

'Pointer
Private Const S_OK As Long = &H0&
Private Const VT_BYREF As Long = &H4000&
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
Private Const FADF_CREATESPAN As Long = &H2000 ' Flag for safearraydestroydata

Private Const MAX_ARRAY_LENGHT As Long = &H7FEFFFFF
Private Const MIN_INDEX As Long = 0 'Option Base { 0 | 1 } 'Dont change!
Private Const NOT_FOUND As Long = MIN_INDEX - 1 'Dont change!
Private Const NOT_SPECIFIED As Long = NOT_FOUND 'Dont change!

Private Const F_FEATURES As Long = 1088 'Dont change!
Private Const ELEMENT_SIZE As Long = 4& 'Dont change!
Private Const ELEMENT_TYPE As Integer = 9 'vba.VbVarType.vbObject 'Dont change!

Private Const INSERTIONSORT_THRESHOLD As Long = 24 'Sould be small

Public Enum SortOrder
    Descending = -1
    Ascending = 1
End Enum

Private Const DEFAULT_SORTORDER As Long = SortOrder.Ascending

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

Private Type Span
    cDims        As Integer
    fFeatures    As Integer
    cbElements   As Long
    cLocks       As Long
    #If Win64 Then
        pvData       As LongPtr
    #Else
        pvData       As Long
    #End If
    cElements    As Long
    lLbound      As Long
End Type

Private Type ObjectTransfer
    #If Win64 Then
        bit(1 To 8) As Byte
    #Else
        bit(1 To 4) As Byte
    #End If
End Type

Private Type Member
    #If Win64 Then
        Descriptor As LongPtr
    #Else
        Descriptor As Long
    #End If
    SafeArray As SAFEARRAYSTRUCT
    Length As Long
    Internal() As Object
    Index() As Long
End Type
Private This As Member

Public Property Get Stream() As GenericSequence: Set Stream = GenericSequence.Stream(Me): End Property
Public Property Get Instance() As IGeneric: Set Instance = Me: End Property
Public Property Get Elements() As IGenericCollection: Set Elements = Me: End Property
Public Property Get MaxLength() As Long: MaxLength = MAX_ARRAY_LENGHT: End Property

Public Property Get LowerBound() As Long: LowerBound = MIN_INDEX: End Property
Public Property Get IsEmpty() As Boolean: IsEmpty = (This.Length = 0): End Property
Public Property Get Descriptor() As LongPtr: Descriptor = This.Descriptor: End Property
'Public Property Get LengthOf(ByVal Dimension As Long) As Long: Call Catch(SafeArrayGetUBound(ByVal This.Descriptor, Dimension, LengthOf)): End Property
Public Property Get LengthOf(ByVal Dimension As Long) As Long: LengthOf = This.SafeArray.Bounds(This.SafeArray.cDims - Dimension).cElements: End Property
Public Property Get Length() As Long: Length = This.Length: End Property
Public Property Get Features() As Integer: Features = This.SafeArray.fFeatures: End Property
'Public Property Get Size() As LongPtr: Size = Me.SizeOf(This.Length): End Property
'Public Property Get Rank() As Long: Rank = SafeArrayGetDim(ByVal This.Descriptor): End Property
Public Property Get Rank() As Long: Rank = This.SafeArray.cDims: End Property
Public Property Set ElementAt(ByVal Index As Long, ByVal Value As Object): Set This.Internal(Index) = Value: End Property
Public Property Get ElementAt(ByVal Index As Long) As Object: Set ElementAt = This.Internal(Index): End Property

Public Function Build(ParamArray Indices() As Variant) As GenericArray

    Dim newArray As Member
    Dim Index As Long
    Dim i As Long, ub As Long, lb As Long

    If IsMissing(Indices) Then _
        Call GenericError.Build(Me, "Public Function Build").InvalidInput("ParamArray Indices are missing")

    lb = 0&: ub = UBound(Indices)

    With newArray.SafeArray

        ReDim .Bounds(lb To ub)

        .cDims = ub + 1
        .fFeatures = F_FEATURES
        .cbElements = ELEMENT_SIZE
        newArray.Length = 1

        For i = lb To ub
            Index = Indices(i) 'try cast

            If (Index < 0) Then _
                Call GenericError.Build(Me, "Public Function Build").ArgOutOfRange("A zero(one)-based array indexed array has its first item indexed as 0(1)")

            newArray.Length = newArray.Length * Index
            .Bounds(i).cElements = Index
            .Bounds(i).lLbound = MIN_INDEX

        Next i

        newArray.Descriptor = SafeArrayCreateEx(ELEMENT_TYPE, .cDims, .Bounds(lb), ByVal 0&)

        If (newArray.Descriptor = 0) Then _
            Call GenericError.Build(Me, "Public Function Build").InvalidOperation("Descriptor = 0 :: Requested Elements: " & newArray.Length)

        If Not (SafeArrayAccessData(ByVal newArray.Descriptor, .pvData)) = S_OK Then _
            Call GenericError.Build(Me, "Public Function Build").InvalidOperation("SafeArrayAccessData")

        If Not (SafeArrayUnaccessData(ByVal newArray.Descriptor)) = S_OK Then _
            Call GenericError.Build(Me, "Public Function Build").InvalidOperation("SafeArrayUnaccessData")

        If .cDims > 1 Then _
            For i = lb To ub: .Bounds(.cDims - 1 - i).cElements = Indices(i): Next 'Roll back
    End With

    'Everything is fine, load internal
    Call CopyMemory(Dst:=ByVal VarPtrArray(newArray.Internal), Src:=ByVal VarPtr(newArray.Descriptor), Length:=POINTERSIZE)

    Set Build = New GenericArray
    Call System.Inject(Build, VarPtr(newArray), LenB(newArray))
    
End Function

Public Function BuildFrom(ByVal Sequence As IGenericSequence, Optional ByVal InitialSize As Long = 16) As GenericArray

    Dim newArray As Member
    Dim i As Long
    Dim Collection As IGenericCollection
    '@Ignore VariableNotAssigned
    Dim Element As Object
    
    If (Sequence Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function BuildFrom").IsNothing("Sequence")
    
    If (InitialSize < 0) Then _
        Call GenericError.Build(Me, "Public Function BuildFrom").InvalidInput("(InitialSize < 0)")
    
    Set BuildFrom = GenericArray 'Avoid null
    
    If TypeOf Sequence Is IGenericCollection Then
        Set Collection = Sequence ' Dont use Collection.ToArray -> may cause circular reference
        
        If (Collection.Count <= 0) Then _
            Exit Function
        
        InitialSize = Collection.Count
        
    End If
    
    With Sequence.Iterator
        If .HasNext(Element) Then
            
            newArray.SafeArray.cDims = 1
            newArray.SafeArray.fFeatures = F_FEATURES
            newArray.SafeArray.cbElements = ELEMENT_SIZE
            newArray.Length = InitialSize
            ReDim newArray.Internal(newArray.Length - 1)
            
            Do
                
                If (i = newArray.Length) Then
                    newArray.Length = newArray.Length * 2
                    ReDim Preserve newArray.Internal(newArray.Length)
                End If
                
                Set newArray.Internal(i) = Element
                i = i + 1
            
            Loop While .HasNext(Element)
            
            If (newArray.Length > i) Then _
                ReDim Preserve newArray.Internal(i - 1)
            newArray.Length = i
            
            ReDim newArray.SafeArray.Bounds(0)
            
            newArray.SafeArray.Bounds(0).lLbound = MIN_INDEX
            newArray.SafeArray.Bounds(0).cElements = i
            
            Call CopyMemory(Dst:=newArray.Descriptor, Src:=ByVal VarPtrArray(newArray.Internal), Length:=POINTERSIZE)
            
            If (newArray.Descriptor = 0) Then _
                Call GenericError.Build(Me, "Public Function BuildFrom").InvalidOperation("Descriptor = 0")
            
            If Not (SafeArrayAccessData(ByVal newArray.Descriptor, newArray.SafeArray.pvData)) = S_OK Then _
                Call GenericError.Build(Me, "Public Function BuildFrom").InvalidOperation("SafeArrayAccessData")
        
         If Not (SafeArrayUnaccessData(ByVal newArray.Descriptor)) = S_OK Then _
            Call GenericError.Build(Me, "Public Function BuildFrom").InvalidOperation("SafeArrayUnaccessData")
        
        End If
    End With
    
    Set BuildFrom = New GenericArray
    Call System.Inject(BuildFrom, VarPtr(newArray), LenB(newArray))
    
End Function

Public Function Of(ParamArray IGenericValues() As Variant) As IGenericCollection
    
    Dim i As Long
    Dim Values As GenericArray
    
    If IsMissing(IGenericValues) Then _
        Call GenericError.Build(Me, "Public Function Of").NoSuchElement("ParamArray IGenericValues are missing")
    
    Set Values = GenericArray.Build(UBound(IGenericValues) + 1)
    
    With Values
        For i = LBound(IGenericValues) To UBound(IGenericValues)
            
            If Not VBA.IsObject(IGenericValues(i)) Then _
                Call GenericError.Build(Me, "Public Function Of").InvalidInput("Datatype {" & TypeName$(IGenericValues(i)) & "} not allowed")
            
            If (IGenericValues(i) Is Nothing) Then _
                Call GenericError.Build(Me, "Public Function Of").InvalidInput("Element is nothing")
            
            If Not TypeOf IGenericValues(i) Is IGeneric Then _
                Call GenericError.Build(Me, "Public Function Of").InvalidInput("Not TypeOf IGenericValues(i) Is IGeneric")
                
            Set .ElementAt(i) = IGenericValues(i)
          
        Next
    End With
    
    Set Of = Values
    
End Function

'' Works and very fast for large matrices
'Public Function BuildFrom(ByRef Matrix() As IGeneric) As GenericArray
'
'    Dim i As Long, FormatBounds As Boolean
'    Dim newArray As Member
'
'    If Not VarType(Matrix) = vba.VbVarType.vbArray + ELEMENT_TYPE Then _
'        Call GenericError.Build(Me, "Public Function BuildWith").InvalidInput("Datatype {" & TypeName$(Matrix) & "} not allowed")
'
'    Call CopyMemory(Dst:=newArray.Descriptor, Src:=ByVal VarPtrArray(Matrix), Length:=POINTERSIZE)
'
'    If (newArray.Descriptor = 0) Then _
'        Call GenericError.Build(Me, "Public Function BuildFrom").InvalidOperation("Descriptor = 0")
'    Call ZeroMemory(Dst:=ByVal VarPtrArray(Matrix), Length:=POINTERSIZE) 'FillMemory(dst:=ByVal VarPtrArray(Matrix), Length:=POINTERSIZE, Fill:=0)
'
'    newArray.Elements = 1
'    Call LoadGenericArrayStructure(newArray.SafeArray, newArray.Descriptor)
'
'    ReDim newArray.Indices(newArray.SafeArray.cDims - 1)
'    For i = LBound(newArray.SafeArray.Bounds) To UBound(newArray.SafeArray.Bounds)
'        If Not newArray.SafeArray.Bounds(i).lLbound = MIN_INDEX Then _
'            FormatBounds = True
'
'        newArray.SafeArray.Bounds(i).lLbound = MIN_INDEX
'        newArray.Elements = newArray.Elements * newArray.SafeArray.Bounds(i).cElements
'    Next
'
'    If FormatBounds Then _
'        Call CopyMemory(Dst:=ByVal This.Descriptor + 16&, Src:=ByVal VarPtr(newArray.SafeArray.Bounds(0&)), Length:=newArray.SafeArray.cDims * 8&)
'
'    Call LoadGenericArray(newArray.Internal, newArray.Descriptor)
'
'    Set BuildFrom = System.CreateInstance(New GenericArray, VarPtr(newArray), LenB(newArray))
'
'End Function

Private Sub LoadGenericArrayStructure(ByRef structure As SAFEARRAYSTRUCT, ByVal SafeArrayPointer As LongPtr)
    
    If (SafeArrayPointer = 0) Then _
        Call GenericError.Build(Me, "Private Sub LoadGenericArrayStructure").InvalidInput("SafeArrayPointer")

    Call CopyMemory(Dst:=structure, Src:=ByVal SafeArrayPointer, Length:=12& + POINTERSIZE)
    
    With structure
        If (.cDims <= 0) Then _
            Call GenericError.Build(Me, "Private Sub LoadGenericArrayStructure").Overflow(".cDims <= 0")
        
        ReDim .Bounds(.cDims - 1)
        Call CopyMemory(Dst:=.Bounds(0&), Src:=ByVal (SafeArrayPointer + 16&), Length:=.cDims * 8&)
    End With

End Sub

Public Function DescriptorOf(ByRef SomeArray As Variant) As LongPtr
    
    Const VT_BYREF As Long = &H4000
    '@Ignore VariableNotAssigned
    Dim vType As Long
    
    If Not (VarType(SomeArray) And VBA.VbVarType.vbArray) = VBA.VbVarType.vbArray Then _
        GenericError.Build(Me, "ArrayPointer").InvalidInput ("Not (IsArray(SomeArray))")

    Call CopyMemory(Dst:=vType, Src:=ByVal VarPtr(SomeArray), Length:=2&)
    Call CopyMemory(Dst:=DescriptorOf, Src:=ByVal VarPtr(SomeArray) + 8&, Length:=POINTERSIZE)
    
    'if SomeArray is defined x() as ...
    If (vType And VT_BYREF) = VT_BYREF Then _
        Call CopyMemory(Dst:=DescriptorOf, Src:=ByVal DescriptorOf, Length:=POINTERSIZE)
   
End Function

Public Sub PutAt(ByVal Value As Object, ParamArray Indices() As Variant)

    Dim i As Long
    
    If (Not This.Index) = -1 Then _
        ReDim This.Index(This.SafeArray.cDims - 1)
    
    For i = 0 To This.SafeArray.cDims - 1: This.Index(i) = Indices(i): Next
        
    If (SafeArrayPutElement(ByVal This.Descriptor, This.Index(0), ByVal ObjPtr(Value)) = DISP_E_BADINDEX) Then _
        Call GenericError.Build(Me, "Public Sub PutAt").InvalidInput("Invalid Index")
    
End Sub

Public Function GetAt(ParamArray Indices() As Variant) As Object
   
    Dim i As Long
    
    If (Not This.Index) = -1 Then _
        ReDim This.Index(This.SafeArray.cDims - 1)

    For i = 0 To This.SafeArray.cDims - 1: This.Index(i) = Indices(i): Next
    
    If (SafeArrayGetElement(ByVal This.Descriptor, This.Index(0), GetAt) = DISP_E_BADINDEX) Then _
        Call GenericError.Build(Me, "Public Function GetAt").InvalidInput("Invalid Index")
   
End Function
'Public Function ExtractElement(ByVal Descriptor As LongPtr, ByRef Indices() As Long, ByVal Out As LongPtr) As Boolean: ExtractElement = S_OK = SafeArrayGetElement(ByVal Descriptor, Indices(0), ByVal Out): End Function
Public Function SizeOf(ByVal Elements As Long) As Long: SizeOf = Elements * This.SafeArray.cbElements: End Function

Public Function AdressOf(Optional ByVal Element As Long = MIN_INDEX) As LongPtr
     
    If (Element >= This.Length) Then _
        Call GenericError.Build(Me, "Public Function AdressOf").InvalidOperation("(Element >= This.Length)")
    
    If (Element < MIN_INDEX) Then _
        Call GenericError.Build(Me, "Public Function AdressOf").InvalidOperation("(Element < MIN_INDEX)")
    
    AdressOf = This.SafeArray.pvData + (Element - MIN_INDEX) * This.SafeArray.cbElements

End Function

Public Function AdressOfIndicies(ParamArray Index() As Variant) As LongPtr
    
    Dim i As Long
    
    If (Not This.Index) = -1 Then _
        ReDim This.Index(This.SafeArray.cDims - 1)
        
    For i = 0 To This.SafeArray.cDims - 1: This.Index(i) = Index(i): Next

    Call Catch(SafeArrayPtrOfIndex(ByVal Descriptor, This.Index(0), AdressOfIndicies))

End Function

Private Sub Override(ByVal Descriptor As LongPtr, ByRef s As SAFEARRAYSTRUCT)
    Call CopyMemory(Dst:=ByVal Descriptor, Src:=ByVal VarPtr(s), Length:=16&)
    Call CopyMemory(Dst:=ByVal Descriptor + 16&, Src:=ByVal VarPtr(s.Bounds(LBound(s.Bounds))), Length:=s.cDims * 8&)
End Sub

'Private Sub Class_Terminate()
'    'Call Catch(SafeArrayDestroy(ByVal .Descriptor))
'End Sub

Public Function Range(Optional ByVal BeginAt As Long = NOT_SPECIFIED, Optional ByVal Count As Long = NOT_SPECIFIED) As IGenericIterator
    
    If (BeginAt = NOT_SPECIFIED) Then _
        BeginAt = MIN_INDEX
        
    If (BeginAt < MIN_INDEX) Then _
        Call GenericError.Build(Me, "Public Function Range").InvalidInput("(BeginAt < MIN_INDEX)")
    
    If (Count = NOT_SPECIFIED) Then _
        Count = This.Length - BeginAt
        
    If (Count > This.Length - BeginAt) Then _
        Call GenericError.Build(Me, "Public Function Range").InvalidInput("(Length > This.Length - BeginAt)")

    Set Range = GenericArrayIterator.Build(Me, BeginAt, Count)
    
End Function

Public Function Slice(ByVal Index As Long) As GenericArray
    
    Dim Overlay As Span
    Dim Column As GenericArray
    
    If (This.SafeArray.cDims <> 2) Then _
        Call GenericError.Build(Me, "Public Function SlizeColumn").NotSupported("Me.Rank != 2")
            
    If (Index < MIN_INDEX) Or (Index > Me.LengthOf(2)) Then _
        Call GenericError.Build(Me, "Public Function SlizeColumn").IndexOutofRange("Index not valid")
    
    If (Me.IsEmpty) Then _
        Call GenericError.Build(Me, "Public Function SlizeColumn").InvalidOperation("Me.IsEmpty")
        
    Set Column = GenericArray.Build(Me.LengthOf(1))
   
    Overlay.cDims = 1
    Overlay.fFeatures = F_FEATURES
    Overlay.cbElements = ELEMENT_SIZE
    Overlay.cElements = Me.LengthOf(1)
    Overlay.lLbound = MIN_INDEX
    Overlay.pvData = Me.AdressOf((Me.LengthOf(1) * Index))
    
    Call Catch(SafeArrayCopyData(ByVal VarPtr(Overlay), ByVal Column.Descriptor))
    
    Set Slice = Column
    
End Function

Public Sub Transpose()
    
    Const ZEROPOINTER As Long = 0
    
    Dim ptrT1 As LongPtr, ptrT2 As LongPtr
    Dim t1() As ObjectTransfer, t2() As ObjectTransfer
    Dim TotalSize As Long
    Dim i As Long, j As Long, m As Long, n As Long
    
    If (Me.IsEmpty) Then _
        Call GenericError.Build(Me, "Public Sub Transpose").InvalidOperation("SafeArray not allocated")
        
    If (This.SafeArray.cDims <> 2) Then _
        Call GenericError.Build(Me, "Public Sub Transpose").InvalidOperation("Rank <> 2")
        
    With This.SafeArray
        m = .Bounds(1).cElements
        n = .Bounds(0).cElements
        
        ReDim t1(1 To m, 1 To n)
        ReDim t2(1 To n, 1 To m)

        ptrT1 = VarPtr(t1(1, 1))
        ptrT2 = VarPtr(t2(1, 1))
        TotalSize = Me.SizeOf(This.Length)
        
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
        
        Call CopyMemory(Dst:=ByVal This.Descriptor + 16&, Src:=ByVal VarPtr(Bounds(0)), Length:=.cDims * 8&)
        .Bounds(0) = Bounds(1): .Bounds(1) = Bounds(0)
     
    End With

End Sub

Public Sub CopyTo( _
                    ByVal Target As GenericArray, _
                    ByVal Index As Long, _
                    Optional ByVal BeginAt As Long = MIN_INDEX, _
                    Optional ByVal Count As Long = NOT_SPECIFIED)

    Dim SourceStruct As Span
    Dim TargetStruct As Span
    
    If (This.Length = 0) Then _
        Exit Sub
        
    If (Count = NOT_SPECIFIED) Then _
        Count = This.Length - BeginAt
    
    If (Target Is Me) Then _
        Call GenericError.Build(Me, "Public Sub CopyTo").InvalidOperation("(Target Is Me)")
        
    If (Target Is Nothing) Then _
        Call GenericError.Build(Me, "Public Sub CopyTo").IsNothing("Target")
        
    If (Count > Target.Length - Index) Then _
        Call GenericError.Build(Me, "Public Sub CopyTo").InvalidInput("(Count > Target.length - Index)")
    
    If (This.Length - BeginAt < Count) Then _
        Call GenericError.Build(Me, "Public Sub CopyTo").InvalidInput("(This.length - BeginAt < Count)")
    
    SourceStruct.cDims = 1
    SourceStruct.fFeatures = F_FEATURES
    SourceStruct.cbElements = ELEMENT_SIZE
    SourceStruct.cElements = Count
    SourceStruct.lLbound = MIN_INDEX
    
    TargetStruct = SourceStruct
    
    SourceStruct.pvData = Me.AdressOf(BeginAt)
    TargetStruct.pvData = Target.AdressOf(Index)
    
    Call Catch(SafeArrayCopyData(ByVal VarPtr(SourceStruct), ByVal VarPtr(TargetStruct)))
    
End Sub

'Haram ;) - 4 legacy
'Public Sub ResizeTo(ByVal newSize As Long)
'
'    Dim bound As SAFEARRAYBOUND
'
'    If Not (Me.Rank = 1) Then _
'        Call GenericError.Build(Me, "Public Sub ReSizeTo").NotSupported("Me.Rank != 1")
'
'''The hard way ;)
''    With This.SafeArray
''        .Bounds(1).cElements = newSize
''        Dim newStruct As SAFEARRAYSTRUCT
''        Dim newDescriptor As Long
''        newDescriptor = SafeArrayCreateEx(vba.VbVarType.vbObject, .cDims, .Bounds(1), ByVal 0&)
''        newStruct = GetArrayStructure(newDescriptor)
''        Call Skynet.CopyData(Dst:=newStruct.pvData, src:=.pvData, Bytes:=This.Size, ClearSource:=True)
''    End With
''
''    With This
''        Call Catch(SafeArrayDestroy(ByVal Descriptor))
''        .Descriptor = newDescriptor
''        .SafeArray = newStruct
''        .Elements = newSize
''        .Size = newSize * .SafeArray.cbElements
''    End With
''
'    bound.cElements = newSize
'    bound.lLbound = MIN_INDEX
'
'    With This
'         Call Catch(SafeArrayRedim(ByVal .Descriptor, bound))
'         Call LoadGenericArrayStructure(.SafeArray, .Descriptor)
'        .Elements = newSize
'    End With
'
'End Sub

Public Function Distinct( _
                            Optional ByVal Predicate As IGenericPredicator, _
                            Optional ByVal Index As Long = MIN_INDEX, _
                            Optional ByVal Length As Long = NOT_SPECIFIED) As Long

    If (Length = NOT_SPECIFIED) Then _
        Length = This.Length - Index
    
'    if (me.Rank != 1)
'    if (Index < min_index or Index > me.Length )
'    if (Length < 0 or Length > me.Length - Index)
    
    Distinct = DistinctInternal(Objects:=This.Internal, LowerBound:=Index, UpperBound:=Length + Index - 1, Predicate:=Predicate)
  
End Function

Public Function BinarySearch( _
                                ByVal Element As Object, _
                                Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER, _
                                Optional ByVal AscendingOrdered As IGenericPredicator, _
                                Optional ByVal Index As Long = MIN_INDEX, _
                                Optional ByVal Length As Long = NOT_SPECIFIED) As Long
    
    If (Length = NOT_SPECIFIED) Then _
        Length = This.Length - Index

'    if (me.Rank != 1)
'    if (Index < min_index or Index > me.Length )
'    if (Length < 0 or Length > me.Length - Index)

    If (AscendingOrdered Is Nothing) Then
        BinarySearch = BinarySearchInternal(Objects:=This.Internal, Element:=Element, LowerBound:=Index, UpperBound:=Length + Index, Direction:=Direction)  ' ExponentialSearchInternal(Objects:=This.Internal, Element:=Element, LowerBound:=Index, UpperBound:=length + Index - 1, Direction:=Direction)
    Else
        BinarySearch = BinarySearchInternal_Delegate(Objects:=This.Internal, Element:=Element, LowerBound:=Index, UpperBound:=Length + Index, AscendingOrdered:=AscendingOrdered, Direction:=Direction)  'ExponentialSearchInternal_AscendingOrdered(Objects:=This.Internal, Element:=Element, LowerBound:=Index, UpperBound:=length + Index - 1, AscendingOrdered:=AscendingOrdered, Direction:=Direction)
    End If
    
End Function

Public Function IndexOf( _
                            ByVal Element As Object, _
                            Optional ByVal Index As Long = MIN_INDEX, _
                            Optional ByVal Length As Long = NOT_SPECIFIED, _
                            Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER, _
                            Optional ByVal Predicate As IGenericPredicator) As Long
  
    If (Length = NOT_SPECIFIED) Then _
        Length = This.Length - Index
        
'    if (me.Rank != 1)
'    if (Index < min_index or Index > me.Length )
'    if (Length < 0 or Length > me.Length - Index)

    If (Direction = Ascending) Then
        IndexOf = IndexOfInternal(Objects:=This.Internal, Element:=Element, LowerBound:=Index, UpperBound:=Length + Index - 1, Predicate:=Predicate)
    Else
        IndexOf = LastIndexOfInternal(Objects:=This.Internal, Element:=Element, LowerBound:=Index, UpperBound:=Length + Index - 1, Predicate:=Predicate)
    End If
    
End Function

Public Function Reverse( _
                            Optional ByVal Index As Long = MIN_INDEX, _
                            Optional ByVal Length As Long = NOT_SPECIFIED) As GenericArray
    
    If (Length = NOT_SPECIFIED) Then _
        Length = This.Length - Index
    
    Call Validate("Public Function Reverse", Me, Index, Length)
    Call ReverseInternal(Objects:=This.Internal, LowerBound:=Index, UpperBound:=Length + Index - 1)
    
    Set Reverse = Me
    
End Function

Public Function Shuffle( _
                            Optional ByVal Index As Long = MIN_INDEX, _
                            Optional ByVal Length As Long = NOT_SPECIFIED) As GenericArray
     
    If (Length = NOT_SPECIFIED) Then _
        Length = This.Length - Index
        
    Call Validate("Public Function Shuffle", Me, Index, Length)
    
    Call ShuffleInternal(Elements:=This.Internal, LowerBound:=Index, UpperBound:=Length + Index - 1)
    Set Shuffle = Me
    
End Function
'todo linearSearch
Public Function SelectKth( _
                            ByVal Kth As Long, _
                            Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER, _
                            Optional ByVal AscendingOrdered As IGenericPredicator, _
                            Optional ByVal Index As Long = MIN_INDEX, _
                            Optional ByVal Length As Long = NOT_SPECIFIED) As Long

    If (Length = NOT_SPECIFIED) Then _
        Length = This.Length - Index
    
    Call Validate("Public Function SelectKth", Me, Index, Length)
    
    If Kth > MIN_INDEX And Kth < Length + MIN_INDEX - 1 Then
        SelectKth = QuickSelectInternal(Keys:=This.Internal, Kth:=Kth, LowerBound:=Index, UpperBound:=Length + MIN_INDEX - 1, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
    Else
        SelectKth = TopElement(Index:=Index, Length:=Length, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
    End If
    
End Function

Public Function SortTop( _
                            ByVal k As Long, _
                            Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER, _
                            Optional ByVal AscendingOrdered As IGenericPredicator, _
                            Optional ByVal Index As Long = MIN_INDEX, _
                            Optional ByVal Length As Long = NOT_SPECIFIED) As GenericArray

    If (Length = NOT_SPECIFIED) Then _
        Length = This.Length - Index

    Call Validate("Public Function SortTop", Me, Index, Length)

    Call SortTopInternal(Keys:=This.Internal, Top:=k, LowerBound:=Index, UpperBound:=Length + Index - 1, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
    Set SortTop = Me

End Function

Public Function SortTopWith( _
                                ByVal k As Long, _
                                ByVal Items As GenericArray, _
                                Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER, _
                                Optional ByVal AscendingOrdered As IGenericPredicator, _
                                Optional ByVal Index As Long = MIN_INDEX, _
                                Optional ByVal Length As Long = NOT_SPECIFIED) As GenericArray

    If (Length = NOT_SPECIFIED) Then _
        Length = This.Length - Index
    
    Call Validate("Public Function SortTopWith", Me, Index, Length)
    
    If Items Is Nothing Then _
        Call GenericError.Build(Me, "Public Function SortTopWith").IsNothing("Items")
    
    If Not (Items.Rank = This.SafeArray.cDims) Then _
        Call GenericError.Build(Me, "Public Function SortTopWith").InvalidInput(" Not (Items.Rank = Me.Rank)")
        
    If (Length > Items.Length) Then _
        Call GenericError.Build(Me, "Public Function SortTopWith").InvalidOperation("(Length<Items.Length)")
    
    Call SortTopWithItemsInternal(Keys:=This.Internal, Items:=Items, Top:=k, LowerBound:=Index, UpperBound:=Length + Index - 1, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
    Set SortTopWith = Me

End Function

Public Function SortPartial( _
                                ByVal BeginAt As Long, _
                                ByVal EndAt As Long, _
                                Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER, _
                                Optional ByVal AscendingOrdered As IGenericPredicator, _
                                Optional ByVal Index As Long = MIN_INDEX, _
                                Optional ByVal Length As Long = NOT_SPECIFIED) As GenericArray
     
    If (Length = NOT_SPECIFIED) Then _
        Length = This.Length - Index
    
    Call Validate("Public Function SortPartial", Me, Index, Length)
    
    Call SortPartialInternal(Keys:=This.Internal, BeginAt:=BeginAt, EndAt:=EndAt, LowerBound:=Index, UpperBound:=Length + Index - 1, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
    Set SortPartial = Me
    
End Function

Public Function SortPartialWith( _
                                    ByVal Items As GenericArray, _
                                    ByVal BeginAt As Long, _
                                    ByVal EndAt As Long, _
                                    Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER, _
                                    Optional ByVal AscendingOrdered As IGenericPredicator, _
                                    Optional ByVal Index As Long = MIN_INDEX, _
                                    Optional ByVal Length As Long = NOT_SPECIFIED) As GenericArray
     
    If (Length = NOT_SPECIFIED) Then _
        Length = This.Length - Index
    
    Call Validate("Public Function SortPartialWith", Me, Index, Length)
    
    If Items Is Nothing Then _
        Call GenericError.Build(Me, "Public Function SortPartialWith").IsNothing("Items")
    
    If Not (Items.Rank = This.SafeArray.cDims) Then _
        Call GenericError.Build(Me, "Public Function SortPartialWith").InvalidInput(" Not (Items.Rank = Me.Rank)")
        
    If (Length > Items.Length) Then _
        Call GenericError.Build(Me, "Public Function SortPartialWith").InvalidOperation("(Length<Items.Length)")
    
    Call SortPartialWithItemsInternal(Keys:=This.Internal, Items:=Items, BeginAt:=BeginAt, EndAt:=EndAt, LowerBound:=Index, UpperBound:=Length + Index - 1, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
        
    Set SortPartialWith = Me
    
End Function

Public Function Sort( _
                        Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER, _
                        Optional ByVal AscendingOrdered As IGenericPredicator, _
                        Optional ByVal Index As Long = MIN_INDEX, _
                        Optional ByVal Length As Long = NOT_SPECIFIED) As GenericArray
     
    If (Length = NOT_SPECIFIED) Then _
        Length = This.Length - Index
        
    Call Validate("Public Function Sort", Me, Index, Length)
    
    Call SortInternal(Keys:=This.Internal, LowerBound:=Index, UpperBound:=Length + Index - 1, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
    Set Sort = Me
    
End Function

Public Function SortWith( _
                            ByVal Items As GenericArray, _
                            Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER, _
                            Optional ByVal AscendingOrdered As IGenericPredicator, _
                            Optional ByVal Index As Long = MIN_INDEX, _
                            Optional ByVal Length As Long = NOT_SPECIFIED) As GenericArray
    
    If (Length = NOT_SPECIFIED) Then _
        Length = This.Length - Index
    
    Call Validate("Public Function SortWith", Me, Index, Length)
    
    If Items Is Nothing Then _
        Call GenericError.Build(Me, "Public Function SortWith").IsNothing("Items")
    
    If Not (Items.Rank = This.SafeArray.cDims) Then _
        Call GenericError.Build(Me, "Public Function SortWith").InvalidInput(" Not (Items.Rank = Me.Rank)")
        
    If (Length > Items.Length) Then _
        Call GenericError.Build(Me, "Public Function SortWith").InvalidOperation("(Length<Items.Length)")
    
    Call SortWithItemsInternal(Keys:=This.Internal, Items:=Items, LowerBound:=Index, UpperBound:=Length + Index - 1, AscendingOrdered:=AscendingOrdered, Direction:=Direction)
    Set SortWith = Me
    
End Function

Private Sub Validate(ByRef Caller As String, ByVal Target As GenericArray, ByVal Index As Long, ByVal Count As Long)
    
    If (Target Is Nothing) Then _
        Call GenericError.Build(Me, Caller).InvalidInput("Target")
        
    With Target
        If (.IsEmpty) Then _
            Call GenericError.Build(Me, Caller).InvalidOperation("(IsEmpty)")
        
        If Not .Rank = 1 Then _
            Call GenericError.Build(Me, Caller).NotSupported("Rank != 1")
            
        If (Index < MIN_INDEX Or Index > .Length) Then _
            Call GenericError.Build(Me, Caller).ArgOutOfRange("(Index < MIN_INDEX Or Index > Length)")
            
        If (Count < 0 Or Count > .Length - Index) Then _
            Call GenericError.Build(Me, Caller).ArgOutOfRange("(Count < 0 Or Count > .Length - Index) ")
    
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

Private Sub FillInternal( _
                            ByRef Objects() As Object, _
                            ByVal Element As Object, _
                            ByVal LowerBound As Long, _
                            ByVal UpperBound As Long)
                        
    Dim i As Long
    For i = LowerBound To UpperBound: Set Objects(i) = Element: Next
    
End Sub
'TODO Exceptions
Private Function TopElement( _
                                ByVal Index As Long, _
                                ByVal Length As Long, _
                                Optional ByVal AscendingOrdered As IGenericPredicator, _
                                Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER) As Long
    
    '@Ignore VariableNotAssigned
    Dim BestSoFar As Object, Candidate As Object, Pivot As IGenericValue
    Dim i As Long, x As Span, Ptr As Long
    
    If (This.Length = 0) Then _
        Exit Function
        
    If (This.Length < Length) Then _
        Exit Function
    
    x.pvData = Me.AdressOf(Index)
    x.fFeatures = This.SafeArray.fFeatures
    x.cbElements = This.SafeArray.cbElements
    x.cDims = 1
    x.cElements = Length
    
    Ptr = VarPtr(x)
     
    If (AscendingOrdered Is Nothing) Then
        If (Direction = Ascending) Then
        
            Do
                If (SafeArrayGetElement(ByVal Ptr, 0&, ByVal Candidate) = DISP_E_BADINDEX) Then _
                    Call GenericError.Build(Me, "Private Function Private Function TopElement").InvalidInput("Invalid Index")
                
                Set Pivot = Candidate
                
                If Not (Pivot Is Nothing) Then
                    If Not (BestSoFar Is Nothing) Then
                        If Pivot.IsLessThan(BestSoFar) Then
                            Set BestSoFar = Pivot
                            Index = i
                        End If
                    Else
                        Set BestSoFar = Pivot
                    End If
                End If
        
                x.pvData = x.pvData + ELEMENT_SIZE
                i = i + 1
            Loop While (i < x.cElements)
        
        Else
            
            Do
                If (SafeArrayGetElement(ByVal Ptr, 0&, ByVal Candidate) = DISP_E_BADINDEX) Then _
                    Call GenericError.Build(Me, "Private Function Private Function TopElement").InvalidInput("Invalid Index")
                
                Set Pivot = Candidate
                
                If Not (Pivot Is Nothing) Then
                    If Not (BestSoFar Is Nothing) Then
                        If Pivot.IsGreaterThan(BestSoFar) Then
                            Set BestSoFar = Pivot
                            Index = i
                        End If
                    Else
                        Set BestSoFar = Pivot
                    End If
                End If
        
                x.pvData = x.pvData + ELEMENT_SIZE
                i = i + 1
            Loop While (i < x.cElements)
        
        End If
        
    Else
  
        With AscendingOrdered
            If (Direction = Ascending) Then
            
                Do
                    If (SafeArrayGetElement(ByVal Ptr, 0&, ByVal Candidate) = DISP_E_BADINDEX) Then _
                        Call GenericError.Build(Me, "Private Function Private Function TopElement").InvalidInput("Invalid Index")
                    
                    If .IsTrue(BestSoFar, Candidate) Then
                        Set BestSoFar = Candidate
                        Index = i
                    End If
                      
                    x.pvData = x.pvData + ELEMENT_SIZE
                    i = i + 1
                Loop While (i < x.cElements)
                
            Else
                
                Do
                    If (SafeArrayGetElement(ByVal Ptr, 0&, ByVal Candidate) = DISP_E_BADINDEX) Then _
                        Call GenericError.Build(Me, "Private Function Private Function TopElement").InvalidInput("Invalid Index")
                    
                    If .IsTrue(Candidate, BestSoFar) Then
                        Set BestSoFar = Candidate
                        Index = i
                    End If
                      
                    x.pvData = x.pvData + ELEMENT_SIZE
                    i = i + 1
                Loop While (i < x.cElements)
            
            End If
        End With
        
    End If
 
    TopElement = Index
    
End Function

Private Function DistinctInternal( _
                                    ByRef Objects() As Object, _
                                    ByVal LowerBound As Long, _
                                    ByVal UpperBound As Long, _
                                    Optional ByVal Predicate As IGenericPredicator) As Long
    
    Dim Casted As IGenericValue
    Dim uniqueCount As Long
    Dim i As Long
    
    uniqueCount = LowerBound + 1 'first element doesnt need to be shifted
     
    If (UpperBound - LowerBound >= 1) Then
        If (Predicate Is Nothing) Then
            'Find first dup element
            For i = uniqueCount To UpperBound
                
                Set Casted = Objects(uniqueCount - 1) 'Try cast
                
                If (Casted Is Objects(uniqueCount)) Then _
                    Exit For
                
                If Not (Casted Is Nothing) Then _
                    If (Casted.Equals(Objects(uniqueCount))) Then _
                        Exit For
                uniqueCount = uniqueCount + 1
            Next
            
            For i = uniqueCount + 1 To UpperBound 'Shift unique elements
                Set Casted = Objects(i - 1)
                If Not (Casted Is Objects(i)) Then
                    If Not (Casted Is Nothing) Then
                        If Not (Casted.Equals(Objects(i))) Then
                            Set Objects(uniqueCount) = Objects(i)
                            uniqueCount = uniqueCount + 1
                        End If
                    Else
                        Set Objects(uniqueCount) = Objects(i)
                        uniqueCount = uniqueCount + 1
                    End If
                End If
            Next
        
        Else
        
            With Predicate 'Find first dup element
                For i = uniqueCount To UpperBound
                    If (.IsTrue(Objects(uniqueCount - 1), Objects(uniqueCount))) Then _
                        Exit For
                    uniqueCount = uniqueCount + 1
                Next
                
                For i = uniqueCount + 1 To UpperBound 'Shift unique elements
                    If Not (.IsTrue(Objects(i - 1), Objects(i))) Then _
                        Set Objects(uniqueCount) = Objects(i): _
                        uniqueCount = uniqueCount + 1
                Next
            End With
            
            Call FillInternal(Objects:=Objects, Element:=Nothing, LowerBound:=uniqueCount, UpperBound:=UpperBound)
        
        End If
    End If
    
    DistinctInternal = uniqueCount
    
End Function

'Exponential search extends binary search to unbounded lists.
'It starts by finding the first element with an index that is both a power of two and greater than the target value.
'Afterwards, it sets that index as the upper bound, and switches to binary search.
Private Function ExponentialSearchInternal( _
                                            ByRef Objects() As Object, _
                                            ByVal Element As IGenericValue, _
                                            ByVal LowerBound As Long, _
                                            ByVal UpperBound As Long, _
                                            Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER) As Long
                                    
    Dim i As Long
    
    i = LowerBound + 1 '2^0
  
    If (i < UpperBound) Then

        With Element
            If (Direction = Ascending) Then
            
                Do
                    If .IsLessThan(Objects(i)) Then _
                        Exit Do
                    LowerBound = i
                    i = LowerBound * 2 '2^i
                Loop While (i < UpperBound)
                
            Else
                
                Do
                    If .IsGreaterThan(Objects(i)) Then _
                        Exit Do
                    LowerBound = i
                    i = LowerBound * 2 '2^i
                Loop While (i < UpperBound)
            
            End If
        End With
        
    End If
    
    'LowerBound=i/2 When the fractional part is exactly 0.5, Clng,CInt always round it to the nearest even number. Eg: 0.5->0 / 1.5->2
    ExponentialSearchInternal = BinarySearchInternal(Objects:=Objects, Element:=Element, LowerBound:=LowerBound, UpperBound:=VBA.IIf(i < UpperBound, i, UpperBound), Direction:=Direction)

End Function

Private Function BinarySearchInternal( _
                                        ByRef Objects() As Object, _
                                        ByVal Element As IGenericValue, _
                                        ByVal LowerBound As Long, _
                                        ByVal UpperBound As Long, _
                                        Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER) As Long
    
    Dim i As Long
    
    With Element
        If (Direction = Ascending) Then
        
            Do While (LowerBound < UpperBound)
            
                i = LowerBound + ((UpperBound - LowerBound) \ 2) '! truncate the floats
                
                'Leftmost
                If .IsGreaterThan(Objects(i)) Then
                    LowerBound = i + 1
                Else
                    UpperBound = i
                End If
                
            Loop
            
        Else
            
            Do While (LowerBound < UpperBound)
            
                i = LowerBound + ((UpperBound - LowerBound) \ 2) '! truncate the floats
                
                If .IsLessThan(Objects(i)) Then
                    LowerBound = i + 1
                Else
                    UpperBound = i
                End If
                
            Loop
        
        End If
    End With

    BinarySearchInternal = LowerBound
    
End Function

Private Function BinarySearchInternal_Delegate( _
                                                ByRef Objects() As Object, _
                                                ByVal Element As Object, _
                                                ByVal LowerBound As Long, _
                                                ByVal UpperBound As Long, _
                                                ByVal AscendingOrdered As IGenericPredicator, _
                                                Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER) As Long
                                                
    Dim i As Long
    
    With AscendingOrdered
        If (Direction = Ascending) Then
        
            Do While (LowerBound < UpperBound)
                i = LowerBound + ((UpperBound - LowerBound) \ 2) '! truncate the floats
                
                'Rightmost
'                If Not .IsTrue(Element, Objects(i)) Then
'                    LowerBound = i + 1
'                Else
'                    UpperBound = i
'                End If
                
                'Rightmost
'                If .IsTrue(Element, Objects(i)) Then
'                    UpperBound = i
'                Else
'                    LowerBound = i + 1
'                End If

                'Leftmost
                If .IsTrue(Objects(i), Element) Then
                    LowerBound = i + 1
                Else
                    UpperBound = i
                End If
                
            Loop
            
        Else
        
            Do While (LowerBound < UpperBound)
                i = LowerBound + ((UpperBound - LowerBound) \ 2) '! truncate the floats

                If .IsTrue(Element, Objects(i)) Then
                    LowerBound = i + 1
                Else
                    UpperBound = i
                End If
            Loop
            
        End If
    End With

    BinarySearchInternal_Delegate = LowerBound
    
End Function

'Exponential search extends binary search to unbounded lists.
'It starts by finding the first element with an index that is both a power of two and greater than the target value.
'Afterwards, it sets that index as the upper bound, and switches to binary search.
Private Function ExponentialSearchInternal_Delegate( _
                                                        ByRef Objects() As Object, _
                                                        ByVal Element As Object, _
                                                        ByVal LowerBound As Long, _
                                                        ByVal UpperBound As Long, _
                                                        ByVal AscendingOrdered As IGenericPredicator, _
                                                        Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER) As Long
                                    
    Dim i As Long
 
    i = LowerBound + 1 '2^0
    
    If (i < UpperBound) Then
    
        With AscendingOrdered
            If (Direction = Ascending) Then
            
                Do
                    If .IsTrue(Element, Objects(i)) Then _
                        Exit Do
                    LowerBound = i
                    i = LowerBound * 2 '2^i
                Loop While (i < UpperBound)
                
            Else
            
                Do
                    If .IsTrue(Objects(i), Element) Then _
                        Exit Do
                    LowerBound = i
                    i = LowerBound * 2 '2^i
                Loop While (i < UpperBound)
                
            End If
        End With
        
    End If
    
    'LowerBound=i/2 When the fractional part is exactly 0.5, Clng,CInt always round it to the nearest even number. Eg: 0.5->0 / 1.5->2
    ExponentialSearchInternal_Delegate = BinarySearchInternal_Delegate(Objects:=Objects, Element:=Element, LowerBound:=LowerBound, UpperBound:=VBA.IIf(i < UpperBound, i, UpperBound), Direction:=Direction, AscendingOrdered:=AscendingOrdered)

End Function

'Returns the index of the first occurrence of a given value in a range of
'this list. It does a linear, O(n) search. The list is searched forwards, starting at index
'and upto count number of elements. The
'elements of the list are compared to the given value using the
'Object.Equals method.

Private Function IndexOfInternal( _
                                    ByRef Objects() As Object, _
                                    ByVal Element As Object, _
                                    ByVal LowerBound As Long, _
                                    ByVal UpperBound As Long, _
                                    Optional ByVal Predicate As IGenericPredicator) As Long
    
    Dim Casted As IGenericValue
    
    If (Predicate Is Nothing) Then
        If Not (Element Is Nothing) Then
            If TypeOf Element Is IGenericValue Then
            
                Set Casted = Element 'Try cast
                With Casted
                    For IndexOfInternal = LowerBound To UpperBound
                        If Not (Objects(IndexOfInternal) Is Nothing) Then _
                            If .Equals(Objects(IndexOfInternal)) Then _
                                Exit Function
                    Next
                End With
                
                IndexOfInternal = NOT_FOUND '!
                Exit Function
            End If
        
            For IndexOfInternal = LowerBound To UpperBound
                If (Objects(IndexOfInternal) Is Element) Then _
                    Exit Function
            Next
            
        End If
        
    Else
    
        With Predicate
            For IndexOfInternal = LowerBound To UpperBound
                If .IsTrue(Objects(IndexOfInternal), Element) Then _
                    Exit Function
            Next
        End With
        
    End If
    
    IndexOfInternal = NOT_FOUND '!
    
End Function

Private Function LastIndexOfInternal( _
                                        ByRef Objects() As Object, _
                                        ByVal Element As Object, _
                                        ByVal LowerBound As Long, _
                                        ByVal UpperBound As Long, _
                                        Optional ByVal Predicate As IGenericPredicator) As Long
    
    Dim Casted As IGenericValue
    
    If (Predicate Is Nothing) Then
        If Not (Element Is Nothing) Then
            If TypeOf Element Is IGenericValue Then
            
                Set Casted = Element 'Try cast
                With Casted
                    For LastIndexOfInternal = UpperBound To LowerBound Step -1
                        If Not (Objects(LastIndexOfInternal) Is Nothing) Then _
                            If .Equals(Objects(LastIndexOfInternal)) Then _
                                Exit Function
                    Next
                End With
                
                LastIndexOfInternal = NOT_FOUND '!
                Exit Function
            End If
        
            For LastIndexOfInternal = UpperBound To LowerBound Step -1
                If (Objects(LastIndexOfInternal) Is Element) Then _
                    Exit Function
            Next
            
        End If
        
    Else
    
        With Predicate
            For LastIndexOfInternal = UpperBound To LowerBound Step -1
                If .IsTrue(Objects(LastIndexOfInternal), Element) Then _
                    Exit Function
            Next
        End With
        
    End If
    
    LastIndexOfInternal = NOT_FOUND '!
    
End Function
'Reverses the elements in a range of this list. Following a call to this
'method, an element in the range given by index and count
'which was previously located at index i will now be located at
'index index + (index + count - i - 1).
Private Sub ReverseInternal( _
                            ByRef Objects() As Object, _
                            ByVal LowerBound As Long, _
                            ByVal UpperBound As Long)

    Do While (LowerBound < UpperBound): Call SwapInternal(Objects, LowerBound, UpperBound): LowerBound = LowerBound + 1: UpperBound = UpperBound - 1: Loop 'While InterlockedIncrement(i) < InterlockedDecrement(j)

End Sub

'Sorting
'FloorLog2*2
Private Function Depthlimit(ByVal n As Long) As Long
    
    Const LOG2 As Double = 0.693147180559945   'TODO Is this comment still valid? => Dont change!
    
    Depthlimit = VBA.Int(VBA.Log(n) / LOG2)
    
'    Do While n > 1
'        i = i + 1
'        n = (n \ 2)
'    Loop
    
    'Depthlimit = 2 * i

End Function

Private Function QuickSelectInternal( _
                                        ByRef Keys() As Object, _
                                        ByVal Kth As Long, _
                                        ByVal LowerBound As Long, _
                                        ByVal UpperBound As Long, _
                                        Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER, _
                                        Optional ByVal AscendingOrdered As IGenericPredicator) As Long
        
    If (UpperBound - LowerBound + 1) < 2 Then _
        Exit Function

    QuickSelectInternal = QuickSelect(Keys:=Keys, Kth:=Kth, Left:=LowerBound, Right:=UpperBound, Direction:=Direction, AscendingOrdered:=AscendingOrdered)

End Function

Private Sub SortPartialInternal( _
                                    ByRef Keys() As Object, _
                                    ByVal BeginAt As Long, _
                                    ByVal EndAt As Long, _
                                    ByVal LowerBound As Long, _
                                    ByVal UpperBound As Long, _
                                    Optional ByVal AscendingOrdered As IGenericPredicator, _
                                    Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER)
    
    If (UpperBound - LowerBound + 1) < 2 Then _
        Exit Sub
        
    If (AscendingOrdered Is Nothing) Then
        Call PartialQuickSort(Keys:=Keys, Left:=LowerBound, Right:=UpperBound, Direction:=Direction, BeginAt:=BeginAt, EndAt:=EndAt)  'Try cast to IGenericValue
    Else
        Call PartialQuickSort_Delegate(Keys:=Keys, Left:=LowerBound, Right:=UpperBound, Direction:=Direction, AscendingOrdered:=AscendingOrdered, BeginAt:=BeginAt, EndAt:=EndAt)
    End If

End Sub

Private Sub SortPartialWithItemsInternal( _
                                            ByRef Keys() As Object, _
                                            ByVal Items As GenericArray, _
                                            ByVal BeginAt As Long, _
                                            ByVal EndAt As Long, _
                                            ByVal LowerBound As Long, _
                                            ByVal UpperBound As Long, _
                                            Optional ByVal AscendingOrdered As IGenericPredicator, _
                                            Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER)
    
    If (UpperBound - LowerBound + 1) < 2 Then _
        Exit Sub
        
    If (AscendingOrdered Is Nothing) Then
        Call PartialQuickSortWithItems(Keys:=Keys, Items:=Items, Left:=LowerBound, Right:=UpperBound, Direction:=Direction, BeginAt:=BeginAt, EndAt:=EndAt) 'Try cast to IGenericValue
    '@Ignore EmptyElseBlock
    Else
        'TODO
    End If

End Sub

Private Sub SortInternal( _
                            ByRef Keys() As Object, _
                            ByVal LowerBound As Long, _
                            ByVal UpperBound As Long, _
                            Optional ByVal AscendingOrdered As IGenericPredicator, _
                            Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER)
    
    If (UpperBound - LowerBound + 1) < 2 Then _
        Exit Sub
        
    If (AscendingOrdered Is Nothing) Then
        Call IntroSort(Keys:=Keys, Left:=LowerBound, Right:=UpperBound, Depthlimit:=Depthlimit(UpperBound), Direction:=Direction)   'Try cast to IGenericValue
    Else
        Call IntroSort_Delegate(Keys:=Keys, Left:=LowerBound, Right:=UpperBound, Depthlimit:=Depthlimit(UpperBound), Direction:=Direction, AscendingOrdered:=AscendingOrdered)
    End If

End Sub

Private Sub SortWithItemsInternal( _
                                    ByRef Keys() As Object, _
                                    ByVal Items As GenericArray, _
                                    ByVal LowerBound As Long, _
                                    ByVal UpperBound As Long, _
                                    Optional ByVal AscendingOrdered As IGenericPredicator, _
                                    Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER)

    If (UpperBound - LowerBound + 1) < 2 Then _
        Exit Sub
        
    If (AscendingOrdered Is Nothing) Then
        Call IntroSortWithItems(Keys:=Keys, Items:=Items, Left:=LowerBound, Right:=UpperBound, Depthlimit:=Depthlimit(UpperBound), Direction:=Direction)  'Try cast to IGenericValue
    Else
        Call IntroSortWithItems_Delegate(Keys:=Keys, Items:=Items, Left:=LowerBound, Right:=UpperBound, Depthlimit:=Depthlimit(UpperBound), Direction:=Direction, AscendingOrdered:=AscendingOrdered)
    End If

End Sub

Private Sub SortTopInternal( _
                                ByRef Keys() As Object, _
                                ByVal Top As Long, _
                                ByVal LowerBound As Long, _
                                ByVal UpperBound As Long, _
                                Optional ByVal AscendingOrdered As IGenericPredicator, _
                                Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER)

    If (UpperBound - LowerBound + 1) < 2 Then _
        Exit Sub

    If (AscendingOrdered Is Nothing) Then
        Call PartialIntroSort(Keys:=Keys, Left:=LowerBound, Right:=UpperBound, Depthlimit:=Depthlimit(UpperBound), Direction:=Direction, Top:=Top) 'Try cast to IGenericValue
    Else
        Call PartialIntroSort_Delegate(Keys:=Keys, Left:=LowerBound, Right:=UpperBound, Depthlimit:=Depthlimit(UpperBound), Direction:=Direction, AscendingOrdered:=AscendingOrdered, Top:=Top)
    End If

End Sub

Private Sub SortTopWithItemsInternal( _
                                        ByRef Keys() As Object, _
                                        ByVal Items As GenericArray, _
                                        ByVal Top As Long, _
                                        ByVal LowerBound As Long, _
                                        ByVal UpperBound As Long, _
                                        Optional ByVal AscendingOrdered As IGenericPredicator, _
                                        Optional ByVal Direction As SortOrder = DEFAULT_SORTORDER)
    
    If (UpperBound - LowerBound + 1) < 2 Then _
        Exit Sub
        
    If (AscendingOrdered Is Nothing) Then
        Call PartialIntroSortWithItems(Keys:=Keys, Items:=Items, Left:=LowerBound, Right:=UpperBound, Depthlimit:=Depthlimit(UpperBound), Direction:=Direction, Top:=Top) 'Try cast to IGenericValue
    Else
        Call PartialIntroSortWithItems_Delegate(Keys:=Keys, Items:=Items, Left:=LowerBound, Right:=UpperBound, Depthlimit:=Depthlimit(UpperBound), Direction:=Direction, AscendingOrdered:=AscendingOrdered, Top:=Top)
    End If

End Sub

Public Sub ShuffleInternal( _
                            ByRef Elements() As Object, _
                            ByVal LowerBound As Long, _
                            ByVal UpperBound As Long)

    Dim i As Long
    
    If (UpperBound - LowerBound + 1) < 2 Then _
        Exit Sub
        
    Randomize VBA.Timer
    
    For i = UpperBound To LowerBound Step -1: Call SwapInternal(Elements, i, VBA.Int((UpperBound - LowerBound + 1) * Rnd + LowerBound)): Next

End Sub

Private Function QuickSelect( _
                                ByRef Keys() As Object, _
                                ByVal Kth As Long, _
                                ByVal Left As Long, _
                                ByVal Right As Long, _
                                ByVal Direction As SortOrder, _
                                Optional ByVal AscendingOrdered As IGenericPredicator) As Long
    
    
    Dim Pivot As IGenericValue, tmp As Object
    Dim i As Long, j As Long, t As Long

    If (Left > Right) Then _
        Exit Function
        
    ReDim Map(Left To Right) As Long
    For i = Left To Right: Map(i) = i: Next
  
    If (Direction = Ascending) Then
    
        Do
        
            i = Left: j = Right
            Set Pivot = Keys(Map(i + ((j - i) \ 2))) 'try to cast
            
            Do
    
                If Not (Pivot Is Nothing) Then
                
                    Do
                        Set tmp = Keys(Map(i))
                        If tmp Is Nothing Then _
                            Exit Do
    
                        If Not Pivot.IsGreaterThan(tmp) Then _
                            Exit Do
                        i = i + 1
                    Loop
    
                    Do
                        Set tmp = Keys(Map(j))
                        If Not tmp Is Nothing Then _
                            If Not Pivot.IsLessThan(tmp) Then _
                                Exit Do
                        j = j - 1
                    Loop
    
                Else
                    
                    Do While Not (Keys(Map(i)) Is Nothing): i = i + 1: Loop
                
                End If
    
                If (i > j) Then _
                    Exit Do
    
                If (i < j) Then
                    t = Map(i)
                    Map(i) = Map(j)
                    Map(j) = t
                End If
                
                i = i + 1: j = j - 1
    
            Loop While (i <= j)
        
            If (i <= Kth) Then
                Left = i + 1
            Else
                Right = j - 1
            End If
            
            If (j - Left) <= (Right - i) Then
                If (Left < j) Then _
                    Right = j
                Left = i
            Else
                If (i < Right) Then _
                    Left = i
                Right = j
            End If
            
        Loop While (Left < Right)
    
    Else
        
        Do
    
            i = Left: j = Right
            Set Pivot = Keys(Map(i + ((j - i) \ 2))) 'try to cast
            
            Do
        
                If Not (Pivot Is Nothing) Then
                
                    Do
                        Set tmp = Keys(Map(i))
                        If tmp Is Nothing Then _
                            Exit Do
    
                        If Not Pivot.IsLessThan(tmp) Then _
                            Exit Do
                        i = i + 1
                    Loop
    
                    Do
                        Set tmp = Keys(Map(j))
                        If Not tmp Is Nothing Then _
                            If Not Pivot.IsGreaterThan(tmp) Then _
                                Exit Do
                        j = j - 1
                    Loop
    
                Else
                    
                    Do While Not (Keys(Map(i)) Is Nothing): i = i + 1: Loop
                
                End If
    
                If (i > j) Then _
                    Exit Do
    
                If (i < j) Then
                    t = Map(i)
                    Map(i) = Map(j)
                    Map(j) = t
                End If
    
                i = i + 1: j = j - 1
    
            Loop While (i <= j)
        
            If (i <= Kth) Then
                Left = i + 1
            Else
                Right = j - 1
            End If
            
            If (j - Left) <= (Right - i) Then
                If (Left < j) Then _
                    Right = j
                Left = i
            Else
                If (i < Right) Then _
                    Left = i
                Right = j
            End If
    
        Loop While (Left < Right)
        
    End If
    
    QuickSelect = Map(Kth)
   
End Function

Private Function QuickSelect_Delegate( _
                                        ByRef Keys() As Object, _
                                        ByVal Kth As Long, _
                                        ByVal Left As Long, _
                                        ByVal Right As Long, _
                                        ByVal Direction As SortOrder, _
                                        Optional ByVal AscendingOrdered As IGenericPredicator) As Long
    
    
    Dim Pivot As Object
    Dim i As Long, j As Long, tmp As Long
   
    If (Left > Right) Then _
        Exit Function
        
    ReDim Map(Left To Right) As Long
    For i = Left To Right: Map(i) = i: Next
        
    With AscendingOrdered
        If (Direction = Ascending) Then
        
            Do
            
                i = Left: j = Right
                Set Pivot = Keys(Map(i + ((j - i) \ 2)))
               
                Do
                    
                    Do While .IsTrue(Keys(i), Pivot): i = i + 1: Loop
                    Do While .IsTrue(Pivot, Keys(j)): j = j - 1: Loop
                        
                    If (i > j) Then _
                        Exit Do
        
                    If (i < j) Then
                        tmp = Map(i)
                        Map(i) = Map(j)
                        Map(j) = tmp
                    End If
        
                    i = i + 1: j = j - 1
        
                Loop While (i <= j)
               
                If (i <= Kth) Then
                    Left = i + 1
                Else
                    Right = j - 1
                End If
                
                If (j - Left) <= (Right - i) Then
                    If (Left < j) Then _
                        Right = j
                    Left = i
                Else
                    If (i < Right) Then _
                        Left = i
                    Right = j
                End If
            
            Loop While (Left < Right)
            
        Else
            
            Do
            
                i = Left: j = Right
                Set Pivot = Keys(Map(i + ((j - i) \ 2)))
               
                Do
                    
                    Do While .IsTrue(Pivot, Keys(i)): i = i + 1: Loop
                    Do While .IsTrue(Keys(j), Pivot): j = j - 1: Loop
                        
                    If (i > j) Then _
                        Exit Do
        
                    If (i < j) Then
                        tmp = Map(i)
                        Map(i) = Map(j)
                        Map(j) = tmp
                    End If
        
                    i = i + 1: j = j - 1
        
                Loop While (i <= j)
               
                If (i <= Kth) Then
                    Left = i + 1
                Else
                    Right = j - 1
                End If
                
                If (j - Left) <= (Right - i) Then
                    If (Left < j) Then _
                        Right = j
                    Left = i
                Else
                    If (i < Right) Then _
                        Left = i
                    Right = j
                End If
            
            Loop While (Left < Right)
            
        End If
    End With
    
    QuickSelect_Delegate = Map(Kth)
   
End Function

Private Sub PartialQuickSort( _
                                ByRef Keys() As Object, _
                                ByVal Left As Long, _
                                ByVal Right As Long, _
                                ByVal Direction As SortOrder, _
                                ByVal BeginAt As Long, _
                                ByVal EndAt As Long)

    Dim Pivot As IGenericValue, tmp As Object
    Dim i As Long, j As Long
    
    If (Direction = Ascending) Then
    
        Do While (Left < Right)
        
            i = Left: j = Right
            Set Pivot = Keys(i + ((j - i) \ 2)) 'try to cast
            
            Do
    
                If Not (Pivot Is Nothing) Then

                    Do
                        Set tmp = Keys(i)
                        If tmp Is Nothing Then _
                            Exit Do
    
                        If Not Pivot.IsGreaterThan(tmp) Then _
                            Exit Do
                        i = i + 1
                    Loop
    
                    Do
                        Set tmp = Keys(j)
                        If Not tmp Is Nothing Then _
                            If Not Pivot.IsLessThan(tmp) Then _
                                Exit Do
                        j = j - 1
                    Loop
    
                Else
                    
                    Do While Not (Keys(i) Is Nothing): i = i + 1: Loop
                
                End If
                
    
                If (i > j) Then _
                    Exit Do
    
                If (i < j) Then
                    Set tmp = Keys(i)
                    Set Keys(i) = Keys(j)
                    Set Keys(j) = tmp
                End If
                
                i = i + 1: j = j - 1
    
            Loop While (i <= j)
        
            If (BeginAt >= i) Then
                Left = Left + 1
            ElseIf (EndAt <= j) Then
                Right = j - 1
            End If
            
            If (j - Left) <= (Right - i) Then
                If (Left < j) Then _
                    Call PartialQuickSort(Keys:=Keys, Left:=Left, Right:=j, Direction:=Direction, BeginAt:=BeginAt, EndAt:=EndAt)
                Left = i
            Else
                If (i < Right) Then _
                    Call PartialQuickSort(Keys:=Keys, Left:=i, Right:=Right, Direction:=Direction, BeginAt:=BeginAt, EndAt:=EndAt)
                Right = j
            End If
            
        Loop
   
    Else
        
        Do While (Left < Right)
    
            i = Left: j = Right
            Set Pivot = Keys(i + ((j - i) \ 2)) 'try to cast
            
            Do
        
                If Not (Pivot Is Nothing) Then
                
                    Do
                        Set tmp = Keys(i)
                        If tmp Is Nothing Then _
                            Exit Do
    
                        If Not Pivot.IsLessThan(tmp) Then _
                            Exit Do
                        i = i + 1
                    Loop
    
                    Do
                        Set tmp = Keys(j)
                        If Not tmp Is Nothing Then _
                            If Not Pivot.IsGreaterThan(tmp) Then _
                                Exit Do
                        j = j - 1
                    Loop
                    
                Else
                    
                    Do While Not (Keys(i) Is Nothing): i = i + 1: Loop
                        
                End If
    
                If (i > j) Then _
                    Exit Do
    
                If (i < j) Then
                    Set tmp = Keys(i)
                    Set Keys(i) = Keys(j)
                    Set Keys(j) = tmp
                End If
                
                i = i + 1: j = j - 1
    
            Loop While (i <= j)
        
            If (BeginAt >= i) Then
                Left = Left + 1
            ElseIf (EndAt <= j) Then
                Right = j - 1
            End If
            
            If (j - Left) <= (Right - i) Then
                If (Left < j) Then _
                    Call PartialQuickSort(Keys:=Keys, Left:=Left, Right:=j, Direction:=Direction, BeginAt:=BeginAt, EndAt:=EndAt)
                Left = i
            Else
                If (i < Right) Then _
                    Call PartialQuickSort(Keys:=Keys, Left:=i, Right:=Right, Direction:=Direction, BeginAt:=BeginAt, EndAt:=EndAt)
                Right = j
            End If
            
        Loop
        
    End If
     
End Sub

Private Sub PartialQuickSort_Delegate( _
                                        ByRef Keys() As Object, _
                                        ByVal Left As Long, _
                                        ByVal Right As Long, _
                                        ByVal Direction As SortOrder, _
                                        ByVal AscendingOrdered As IGenericPredicator, _
                                        ByVal BeginAt As Long, _
                                        ByVal EndAt As Long)

    Dim Pivot As Object, tmp As Object
    Dim i As Long, j As Long
    
    With AscendingOrdered
        If (Direction = Ascending) Then
        
            Do While (Left < Right)
            
                i = Left: j = Right
                Set Pivot = Keys(i + ((j - i) \ 2)) 'try to cast
                
                Do
                    
                    Do While .IsTrue(Keys(i), Pivot): i = i + 1: Loop
                    Do While .IsTrue(Pivot, Keys(j)): j = j - 1: Loop
                    
                    If (i > j) Then _
                        Exit Do
        
                    If (i < j) Then
                        Set tmp = Keys(i)
                        Set Keys(i) = Keys(j)
                        Set Keys(j) = tmp
                    End If
                    
                    i = i + 1: j = j - 1
        
                Loop While (i <= j)
            
                If (BeginAt >= i) Then
                    Left = Left + 1
                ElseIf (EndAt <= j) Then
                    Right = j - 1
                End If
                
                If (j - Left) <= (Right - i) Then
                    If (Left < j) Then _
                        Call PartialQuickSort_Delegate(Keys:=Keys, Left:=Left, Right:=j, Direction:=Direction, AscendingOrdered:=AscendingOrdered, BeginAt:=BeginAt, EndAt:=EndAt)
                    Left = i
                Else
                    If (i < Right) Then _
                        Call PartialQuickSort_Delegate(Keys:=Keys, Left:=i, Right:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered, BeginAt:=BeginAt, EndAt:=EndAt)
                    Right = j
                End If
                
            Loop
            
        Else
            
            Do While (Left < Right)
            
                i = Left: j = Right
                Set Pivot = Keys(i + ((j - i) \ 2)) 'try to cast
                
                Do
                    
                    Do While .IsTrue(Pivot, Keys(i)): i = i + 1: Loop
                    Do While .IsTrue(Keys(j), Pivot): j = j - 1: Loop
                    
                    If (i > j) Then _
                        Exit Do
        
                    If (i < j) Then
                        Set tmp = Keys(i)
                        Set Keys(i) = Keys(j)
                        Set Keys(j) = tmp
                    End If
                    
                    i = i + 1: j = j - 1
        
                Loop While (i <= j)
            
                If (BeginAt >= i) Then
                    Left = Left + 1
                ElseIf (EndAt <= j) Then
                    Right = j - 1
                End If
                
                If (j - Left) <= (Right - i) Then
                    If (Left < j) Then _
                        Call PartialQuickSort_Delegate(Keys:=Keys, Left:=Left, Right:=j, Direction:=Direction, AscendingOrdered:=AscendingOrdered, BeginAt:=BeginAt, EndAt:=EndAt)
                    Left = i
                Else
                    If (i < Right) Then _
                        Call PartialQuickSort_Delegate(Keys:=Keys, Left:=i, Right:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered, BeginAt:=BeginAt, EndAt:=EndAt)
                    Right = j
                End If
                
            Loop
        
        End If
    End With
     
End Sub

Private Sub PartialQuickSortWithItems( _
                                        ByRef Keys() As Object, _
                                        ByVal Items As GenericArray, _
                                        ByVal Left As Long, _
                                        ByVal Right As Long, _
                                        ByVal Direction As SortOrder, _
                                        ByVal BeginAt As Long, _
                                        ByVal EndAt As Long)

    Dim Pivot As IGenericValue, tmp As Object
    Dim i As Long, j As Long
    
    With Items
        If (Direction = Ascending) Then
        
            Do While (Left < Right)
            
                i = Left: j = Right
                Set Pivot = Keys(i + ((j - i) \ 2)) 'try to cast
                
                Do
        
                    If Not (Pivot Is Nothing) Then
    
                        Do
                            Set tmp = Keys(i)
                            If tmp Is Nothing Then _
                                Exit Do
        
                            If Not Pivot.IsGreaterThan(tmp) Then _
                                Exit Do
                            i = i + 1
                        Loop
        
                        Do
                            Set tmp = Keys(j)
                            If Not tmp Is Nothing Then _
                                If Not Pivot.IsLessThan(tmp) Then _
                                    Exit Do
                            j = j - 1
                        Loop
        
                    Else
                        
                        Do While Not (Keys(i) Is Nothing): i = i + 1: Loop
                    
                    End If
                    
        
                    If (i > j) Then _
                        Exit Do
        
                    If (i < j) Then
                        Set tmp = Keys(i)
                        Set Keys(i) = Keys(j)
                        Set Keys(j) = tmp
                        
                        Set tmp = .ElementAt(i)
                        Set .ElementAt(i) = .ElementAt(j)
                        Set .ElementAt(j) = tmp
                    End If
                    
                    i = i + 1: j = j - 1
        
                Loop While (i <= j)
            
                If (BeginAt >= i) Then
                    Left = Left + 1
                ElseIf (EndAt <= j) Then
                    Right = j - 1
                End If
                
                If (j - Left) <= (Right - i) Then
                    If (Left < j) Then _
                        Call PartialQuickSortWithItems(Keys:=Keys, Items:=Items, Left:=Left, Right:=j, Direction:=Direction, BeginAt:=BeginAt, EndAt:=EndAt)
                    Left = i
                Else
                    If (i < Right) Then _
                        Call PartialQuickSortWithItems(Keys:=Keys, Items:=Items, Left:=i, Right:=Right, Direction:=Direction, BeginAt:=BeginAt, EndAt:=EndAt)
                    Right = j
                End If
                
            Loop
       
        Else
            
            Do While (Left < Right)
        
                i = Left: j = Right
                Set Pivot = Keys(i + ((j - i) \ 2)) 'try to cast
                
                Do
            
                    If Not (Pivot Is Nothing) Then
                    
                        Do
                            Set tmp = Keys(i)
                            If tmp Is Nothing Then _
                                Exit Do
        
                            If Not Pivot.IsLessThan(tmp) Then _
                                Exit Do
                            i = i + 1
                        Loop
        
                        Do
                            Set tmp = Keys(j)
                            If Not tmp Is Nothing Then _
                                If Not Pivot.IsGreaterThan(tmp) Then _
                                    Exit Do
                            j = j - 1
                        Loop
                        
                    Else
                        
                        Do While Not (Keys(i) Is Nothing): i = i + 1: Loop
                            
                    End If
        
                    If (i > j) Then _
                        Exit Do
        
                    If (i < j) Then
                        Set tmp = Keys(i)
                        Set Keys(i) = Keys(j)
                        Set Keys(j) = tmp
                        
                        Set tmp = .ElementAt(i)
                        Set .ElementAt(i) = .ElementAt(j)
                        Set .ElementAt(j) = tmp
                    End If
                    
                    i = i + 1: j = j - 1
        
                Loop While (i <= j)
            
                If (BeginAt >= i) Then
                    Left = Left + 1
                ElseIf (EndAt <= j) Then
                    Right = j - 1
                End If
                
                If (j - Left) <= (Right - i) Then
                    If (Left < j) Then _
                        Call PartialQuickSortWithItems(Keys:=Keys, Items:=Items, Left:=Left, Right:=j, Direction:=Direction, BeginAt:=BeginAt, EndAt:=EndAt)
                    Left = i
                Else
                    If (i < Right) Then _
                        Call PartialQuickSortWithItems(Keys:=Keys, Items:=Items, Left:=i, Right:=Right, Direction:=Direction, BeginAt:=BeginAt, EndAt:=EndAt)
                    Right = j
                End If
                
            Loop
            
        End If
    End With
     
End Sub

Private Sub PartialIntroSort( _
                                ByRef Keys() As Object, _
                                ByVal Left As Long, _
                                ByVal Right As Long, _
                                ByVal Depthlimit As Long, _
                                ByVal Direction As SortOrder, _
                                ByVal Top As Long)

    Dim PartitionSize As Long
    Dim p As Long
    
    Do While (Right > Left)
       
        PartitionSize = Right - Left + 1

        If (PartitionSize <= INSERTIONSORT_THRESHOLD) Then
            If (PartitionSize > 1) Then
                Select Case PartitionSize
                    Case 2
                        Call InOrderSwap(Keys:=Keys, i:=Left, j:=Right, Direction:=Direction)
                    Case 3
                        Call PivotSwap(Keys:=Keys, i:=Left, j:=Right, Direction:=Direction)
                    Case Else
                        Call BinarySort(Keys:=Keys, Left:=Left, Right:=Right, Direction:=Direction)
                End Select
            End If
            Exit Sub
        End If

        If (Depthlimit = 0) Then _
            Call HeapSort(Keys, Left, Right, Direction): Exit Sub
        
        Depthlimit = Depthlimit - 1
        p = Partition(Keys:=Keys, Left:=Left, Right:=Right, Direction:=Direction)
        
        If Top <= p Then
            Right = p - 1
        ElseIf p - Left > Right - p Then
            Call IntroSort(Keys:=Keys, Left:=p + 1, Right:=Right, Depthlimit:=Depthlimit, Direction:=Direction)
            Right = p - 1
        Else
            Call IntroSort(Keys:=Keys, Left:=Left, Right:=p - 1, Depthlimit:=Depthlimit, Direction:=Direction)
            Left = p + 1
        End If
        
    Loop
    
End Sub

Private Sub IntroSort( _
                        ByRef Keys() As Object, _
                        ByVal Left As Long, _
                        ByVal Right As Long, _
                        ByVal Depthlimit As Long, _
                        ByVal Direction As SortOrder)

    Dim PartitionSize As Long
    Dim p As Long
    
    Do While (Right > Left)
       
        PartitionSize = Right - Left + 1

        If (PartitionSize <= INSERTIONSORT_THRESHOLD) Then
            If (PartitionSize > 1) Then
                Select Case PartitionSize
                    Case 2
                        Call InOrderSwap(Keys:=Keys, i:=Left, j:=Right, Direction:=Direction)
                    Case 3
                        Call PivotSwap(Keys:=Keys, i:=Left, j:=Right, Direction:=Direction)
                    Case Else
                        Call BinarySort(Keys:=Keys, Left:=Left, Right:=Right, Direction:=Direction)
                End Select
            End If
            Exit Sub
        End If

        If (Depthlimit = 0) Then _
            Call HeapSort(Keys, Left, Right, Direction): Exit Sub

        Depthlimit = Depthlimit - 1
        p = Partition(Keys:=Keys, Left:=Left, Right:=Right, Direction:=Direction)
        Call IntroSort(Keys:=Keys, Left:=p + 1, Right:=Right, Depthlimit:=Depthlimit, Direction:=Direction)
        Right = p - 1
    
    Loop

End Sub

Private Function Partition( _
                                ByRef Keys() As Object, _
                                ByVal Left As Long, _
                                ByVal Right As Long, _
                                ByVal Direction As SortOrder) As Long
    
    Dim Pivot As IGenericValue, tmp As Object
    Dim i As Long, j As Long, p As Long
      
    p = PivotSwap(Keys:=Keys, i:=Left, j:=Right, Direction:=Direction)
    Set Pivot = Keys(p) 'try to cast
   
    i = Left: j = Right - 1
    
    Set Keys(p) = Keys(j)
    Set Keys(j) = Pivot
      
    If (Direction = Ascending) Then
    
        Do While (i < j)
            
            If Not (Pivot Is Nothing) Then

                Do
                    i = i + 1
                    Set tmp = Keys(i)
                    If tmp Is Nothing Then _
                        Exit Do

                    If Not Pivot.IsGreaterThan(tmp) Then _
                        Exit Do
                Loop

                Do
                    j = j - 1
                    Set tmp = Keys(j)
                    If Not tmp Is Nothing Then _
                        If Not Pivot.IsLessThan(tmp) Then _
                            Exit Do
                Loop

            Else
                
                Do While (i < Right - 1)
                    i = i + 1
                    If (Keys(i) Is Nothing) Then _
                        Exit Do
                Loop
                
                Do While (j > Left)
                    j = j - 1
                    If Not (Keys(j) Is Nothing) Then _
                        Exit Do
                Loop
            
            End If
        
            If (i >= j) Then _
                Exit Do
        
            Set tmp = Keys(i)
            Set Keys(i) = Keys(j)
            Set Keys(j) = tmp
             
        Loop
      
    Else
        
        Do While (i < j)
            
            If Not (Pivot Is Nothing) Then
            
                Do
                    i = i + 1
                    Set tmp = Keys(i)
                    If tmp Is Nothing Then _
                        Exit Do

                    If Not Pivot.IsLessThan(tmp) Then _
                        Exit Do
                Loop

                Do
                    j = j - 1
                    Set tmp = Keys(j)
                    If Not tmp Is Nothing Then _
                        If Not Pivot.IsGreaterThan(tmp) Then _
                            Exit Do
                Loop
                
            Else
                
                Do While (i < Right - 1)
                    i = i + 1
                    If (Keys(i) Is Nothing) Then _
                        Exit Do
                Loop
                
                Do While (j > Left)
                    j = j - 1
                    If Not (Keys(j) Is Nothing) Then _
                        Exit Do
                Loop
            
            End If
        
            If (i >= j) Then _
                Exit Do
            
            Set tmp = Keys(i)
            Set Keys(i) = Keys(j)
            Set Keys(j) = tmp
             
        Loop
    
    End If
    
    j = (Right - 1)
    
    If Not (i = j) Then
        Set tmp = Keys(i)
        Set Keys(i) = Keys(j)
        Set Keys(j) = tmp
    End If

    Partition = i
        
End Function

Private Sub PartialIntroSortWithItems( _
                                        ByRef Keys() As Object, _
                                        ByVal Items As GenericArray, _
                                        ByVal Left As Long, _
                                        ByVal Right As Long, _
                                        ByVal Depthlimit As Long, _
                                        ByVal Direction As SortOrder, _
                                        ByVal Top As Long)

    Dim PartitionSize As Long
    Dim p As Long
    
    Do While (Right > Left)
       
        PartitionSize = Right - Left + 1

        If (PartitionSize <= INSERTIONSORT_THRESHOLD) Then
            If (PartitionSize > 1) Then
                Select Case PartitionSize
                    Case 2
                        Call InOrderSwapWithItem(Keys:=Keys, Items:=Items, i:=Left, j:=Right, Direction:=Direction)
                    Case 3
                        Call PivotSwapWithItems(Keys:=Keys, Items:=Items, i:=Left, j:=Right, Direction:=Direction)
                    Case Else
                        Call BinarySortWithItems(Keys:=Keys, Items:=Items, Left:=Left, Right:=Right, Direction:=Direction)
                End Select
            End If
            Exit Sub
        End If
        
        If (Depthlimit = 0) Then _
            Call HeapSortWithItems(Keys, Items, Left, Right, Direction): Exit Sub
        
        Depthlimit = Depthlimit - 1
        p = PartitionWithItems(Keys:=Keys, Items:=Items, Left:=Left, Right:=Right, Direction:=Direction)
        
        If (Top <= p) Then
            Right = p - 1
        ElseIf (p - Left > Right - p) Then
            Call IntroSortWithItems(Keys:=Keys, Items:=Items, Left:=p + 1, Right:=Right, Depthlimit:=Depthlimit, Direction:=Direction)
            Right = p - 1
        Else
            Call IntroSortWithItems(Keys:=Keys, Items:=Items, Left:=Left, Right:=p - 1, Depthlimit:=Depthlimit, Direction:=Direction)
            Left = p + 1
        End If
        
    Loop
    
End Sub

Private Sub IntroSortWithItems( _
                                ByRef Keys() As Object, _
                                ByVal Items As GenericArray, _
                                ByVal Left As Long, _
                                ByVal Right As Long, _
                                ByVal Depthlimit As Long, _
                                ByVal Direction As SortOrder)

    Dim PartitionSize As Long
    Dim p As Long
    
    Do While (Right > Left)
       
        PartitionSize = Right - Left + 1

        If (PartitionSize <= INSERTIONSORT_THRESHOLD) Then

            If (PartitionSize > 1) Then
                Select Case PartitionSize
                    Case 2
                        Call InOrderSwapWithItem(Keys:=Keys, Items:=Items, i:=Left, j:=Right, Direction:=Direction)
                    Case 3
                        Call PivotSwapWithItems(Keys:=Keys, Items:=Items, i:=Left, j:=Right, Direction:=Direction)
                    Case Else
                        Call BinarySortWithItems(Keys:=Keys, Items:=Items, Left:=Left, Right:=Right, Direction:=Direction)
                End Select
            End If
            Exit Sub

        End If

        If (Depthlimit = 0) Then _
            Call HeapSortWithItems(Keys, Items, Left, Right, Direction): Exit Sub
        
        Depthlimit = Depthlimit - 1
        p = PartitionWithItems(Keys:=Keys, Items:=Items, Left:=Left, Right:=Right, Direction:=Direction)
        Call IntroSortWithItems(Keys:=Keys, Items:=Items, Left:=p + 1, Right:=Right, Depthlimit:=Depthlimit, Direction:=Direction)
        Right = p - 1
    
    Loop

End Sub

Private Function PartitionWithItems( _
                                        ByRef Keys() As Object, _
                                        ByVal Items As GenericArray, _
                                        ByVal Left As Long, _
                                        ByVal Right As Long, _
                                        ByVal Direction As SortOrder) As Long
                                                
    Dim Pivot As IGenericValue, tmp As Object
    Dim i As Long, j As Long, p As Long
      
    p = PivotSwapWithItems(Keys:=Keys, Items:=Items, i:=Left, j:=Right, Direction:=Direction)
    Set Pivot = Keys(p) 'try to cast
   
    i = Left: j = Right - 1
    
    Set Keys(p) = Keys(j)
    Set Keys(j) = Pivot
    
    With Items
        
        Set tmp = .ElementAt(p)
        Set .ElementAt(p) = .ElementAt(j)
        Set .ElementAt(j) = tmp
        
        If (Direction = Ascending) Then
        
            Do While (i < j)
                
                If Not (Pivot Is Nothing) Then
    
                    Do
                        i = i + 1
                        Set tmp = Keys(i)
                        If tmp Is Nothing Then _
                            Exit Do
    
                        If Not Pivot.IsGreaterThan(tmp) Then _
                            Exit Do
                    Loop
    
                    Do
                        j = j - 1
                        Set tmp = Keys(j)
                        If Not tmp Is Nothing Then _
                            If Not Pivot.IsLessThan(tmp) Then _
                                Exit Do
                    Loop
    
                Else
                    
                    Do While (i < Right - 1)
                        i = i + 1
                        If (Keys(i) Is Nothing) Then _
                            Exit Do
                    Loop
                    
                    Do While (j > Left)
                        j = j - 1
                        If Not (Keys(j) Is Nothing) Then _
                            Exit Do
                    Loop
                
                End If
            
                If (i >= j) Then _
                    Exit Do
            
                Set tmp = Keys(i)
                Set Keys(i) = Keys(j)
                Set Keys(j) = tmp
                
                Set tmp = .ElementAt(i)
                Set .ElementAt(i) = .ElementAt(j)
                Set .ElementAt(j) = tmp
                 
            Loop
          
        Else
            
            Do While (i < j)
                
                If Not (Pivot Is Nothing) Then
                
                    Do
                        i = i + 1
                        Set tmp = Keys(i)
                        If tmp Is Nothing Then _
                            Exit Do
    
                        If Not Pivot.IsLessThan(tmp) Then _
                            Exit Do
                    Loop
    
                    Do
                        j = j - 1
                        Set tmp = Keys(j)
                        If Not tmp Is Nothing Then _
                            If Not Pivot.IsGreaterThan(tmp) Then _
                                Exit Do
                    Loop
                    
                Else
                    
                    Do While (i < Right - 1)
                        i = i + 1
                        If (Keys(i) Is Nothing) Then _
                            Exit Do
                    Loop
                    
                    Do While (j > Left)
                        j = j - 1
                        If Not (Keys(j) Is Nothing) Then _
                            Exit Do
                    Loop
                
                End If
            
                If (i >= j) Then _
                    Exit Do
                
                Set tmp = Keys(i)
                Set Keys(i) = Keys(j)
                Set Keys(j) = tmp
                
                Set tmp = .ElementAt(i)
                Set .ElementAt(i) = .ElementAt(j)
                Set .ElementAt(j) = tmp
                 
            Loop
        
        End If
        
        j = (Right - 1)
        
        If Not (i = j) Then
            Set tmp = Keys(i)
            Set Keys(i) = Keys(j)
            Set Keys(j) = tmp
            
            Set tmp = .ElementAt(i)
            Set .ElementAt(i) = .ElementAt(j)
            Set .ElementAt(j) = tmp
        End If

    End With
    
    PartitionWithItems = i
        
End Function

Private Sub PartialIntroSort_Delegate( _
                                        ByRef Keys() As Object, _
                                        ByVal Left As Long, _
                                        ByVal Right As Long, _
                                        ByVal Depthlimit As Long, _
                                        ByVal Direction As SortOrder, _
                                        ByVal AscendingOrdered As IGenericPredicator, _
                                        ByVal Top As Long)

    Dim PartitionSize As Long
    Dim p As Long
    
    Do While (Right > Left)
        
        PartitionSize = Right - Left + 1

        If (PartitionSize <= INSERTIONSORT_THRESHOLD) Then
            If (PartitionSize > 1) Then
                Select Case PartitionSize
                    Case 2
                        Call InOrderSwap_Delegate(Keys:=Keys, i:=Left, j:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
                    Case 3
                        Call PivotSwap_Delegate(Keys:=Keys, i:=Left, j:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
                    Case Else
                        Call BinarySort_Delegate(Keys:=Keys, Left:=Left, Right:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
                End Select
            End If
            Exit Sub
        End If

        If (Depthlimit = 0) Then _
            Call HeapSort_Delegate(Keys:=Keys, Left:=Left, Right:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered): Exit Sub
        
        Depthlimit = Depthlimit - 1
        p = Partition_Delegate(Keys:=Keys, Left:=Left, Right:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
        
        If Top <= p Then
            Right = p - 1
        ElseIf p - Left > Right - p Then
            Call IntroSort_Delegate(Keys:=Keys, Left:=p + 1, Right:=Right, Depthlimit:=Depthlimit, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
            Right = p - 1
        Else
            Call IntroSort_Delegate(Keys:=Keys, Left:=Left, Right:=p - 1, Depthlimit:=Depthlimit, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
            Left = p + 1
        End If
        
    Loop
    
End Sub

Private Sub IntroSort_Delegate( _
                                ByRef Keys() As Object, _
                                ByVal Left As Long, _
                                ByVal Right As Long, _
                                ByVal Depthlimit As Long, _
                                ByVal Direction As SortOrder, _
                                ByVal AscendingOrdered As IGenericPredicator)

    Dim PartitionSize As Long
    Dim p As Long
    
    Do While (Right > Left)
        
        PartitionSize = Right - Left + 1
            
        If (PartitionSize <= INSERTIONSORT_THRESHOLD) Then
            If (PartitionSize > 1) Then
                Select Case PartitionSize
                    Case 2
                        Call InOrderSwap_Delegate(Keys:=Keys, i:=Left, j:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
                    Case 3
                        Call PivotSwap_Delegate(Keys:=Keys, i:=Left, j:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
                    Case Else
                        Call BinarySort_Delegate(Keys:=Keys, Left:=Left, Right:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
                End Select
            End If
            Exit Sub
        End If
        
        If (Depthlimit = 0) Then _
            Call HeapSort_Delegate(Keys:=Keys, Left:=Left, Right:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered): Exit Sub
        
        Depthlimit = Depthlimit - 1
        p = Partition_Delegate(Keys:=Keys, Left:=Left, Right:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
        Call IntroSort_Delegate(Keys:=Keys, Left:=p + 1, Right:=Right, Depthlimit:=Depthlimit, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
        Right = p - 1
    
    Loop

End Sub

Private Function Partition_Delegate( _
                                        ByRef Keys() As Object, _
                                        ByVal Left As Long, _
                                        ByVal Right As Long, _
                                        ByVal Direction As SortOrder, _
                                        ByVal AscendingOrdered As IGenericPredicator) As Long
    
    Dim Pivot As Object, tmp As Object
    Dim i As Long, j As Long, p As Long
    
    p = PivotSwap_Delegate(Keys:=Keys, i:=Left, j:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
    Set Pivot = Keys(p) 'try to cast
   
    i = Left: j = Right - 1
    
    Set Keys(p) = Keys(j)
    Set Keys(j) = Pivot
    
    With AscendingOrdered
        If Direction = Ascending Then
        
            Do While (i < j)
            
                Do: i = i + 1: Loop While (.IsTrue(Keys(i), Pivot))
                Do: j = j - 1: Loop While (.IsTrue(Pivot, Keys(j)))
                
                If (i >= j) Then _
                    Exit Do
            
                Set tmp = Keys(i)
                Set Keys(i) = Keys(j)
                Set Keys(j) = tmp
                 
            Loop
            
        Else
        
            Do While (i < j)
            
                Do: i = i + 1: Loop While (.IsTrue(Pivot, Keys(i)))
                Do: j = j - 1: Loop While (.IsTrue(Keys(j), Pivot))
                
                If (i >= j) Then _
                    Exit Do
            
                Set tmp = Keys(i)
                Set Keys(i) = Keys(j)
                Set Keys(j) = tmp
                 
            Loop
        
        End If
    End With
    
    j = (Right - 1)
    
    If Not (i = j) Then
        Set tmp = Keys(i)
        Set Keys(i) = Keys(j)
        Set Keys(j) = tmp
    End If
    
    Partition_Delegate = i
        
End Function

Private Sub PartialIntroSortWithItems_Delegate( _
                                                    ByRef Keys() As Object, _
                                                    ByVal Items As GenericArray, _
                                                    ByVal Left As Long, _
                                                    ByVal Right As Long, _
                                                    ByVal Depthlimit As Long, _
                                                    ByVal Direction As SortOrder, _
                                                    ByVal AscendingOrdered As IGenericPredicator, _
                                                    ByVal Top As Long)

    Dim PartitionSize As Long
    Dim p As Long
    
    Do While (Right > Left)
       
        PartitionSize = Right - Left + 1

        If (PartitionSize <= INSERTIONSORT_THRESHOLD) Then
            If (PartitionSize > 1) Then
                Select Case PartitionSize
                    Case 2
                        Call InOrderSwapWithItem_Delegate(Keys:=Keys, Items:=Items, i:=Left, j:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
                    Case 3
                        Call PivotSwapWithItems_Delegate(Keys:=Keys, Items:=Items, i:=Left, j:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
                    Case Else
                        Call BinarySortWithItems_Delegate(Keys:=Keys, Items:=Items, Left:=Left, Right:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
                End Select
            End If
            Exit Sub
        End If
        
        If (Depthlimit = 0) Then _
            Call HeapSortWithItems_Delegate(Keys:=Keys, Items:=Items, Left:=Left, Right:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered): Exit Sub
        
        Depthlimit = Depthlimit - 1
        p = PartitionWithItems_Delegate(Keys:=Keys, Items:=Items, Left:=Left, Right:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
        
        If Top <= p Then
            Right = p - 1
        ElseIf p - Left > Right - p Then
           Call IntroSortWithItems_Delegate(Keys:=Keys, Items:=Items, Left:=p + 1, Right:=Right, Depthlimit:=Depthlimit, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
            Right = p - 1
        Else
            Call IntroSortWithItems_Delegate(Keys:=Keys, Items:=Items, Left:=Left, Right:=p - 1, Depthlimit:=Depthlimit, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
            Left = p + 1
        End If
        
    Loop
    
End Sub

Private Sub IntroSortWithItems_Delegate( _
                                            ByRef Keys() As Object, _
                                            ByVal Items As GenericArray, _
                                            ByVal Left As Long, _
                                            ByVal Right As Long, _
                                            ByVal Depthlimit As Long, _
                                            ByVal Direction As SortOrder, _
                                            ByVal AscendingOrdered As IGenericPredicator)

    Dim PartitionSize As Long
    Dim p As Long
    
    Do While (Right > Left)
        
        PartitionSize = Right - Left + 1
            
        If (PartitionSize <= INSERTIONSORT_THRESHOLD) Then
            If (PartitionSize > 1) Then
                Select Case PartitionSize
                    Case 2
                        Call InOrderSwapWithItem_Delegate(Keys:=Keys, Items:=Items, i:=Left, j:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
                    Case 3
                        Call PivotSwapWithItems_Delegate(Keys:=Keys, Items:=Items, i:=Left, j:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
                    Case Else
                        Call BinarySortWithItems_Delegate(Keys:=Keys, Items:=Items, Left:=Left, Right:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
                End Select
            End If
            Exit Sub
        End If
        
        If (Depthlimit = 0) Then _
            Call HeapSortWithItems_Delegate(Keys:=Keys, Items:=Items, Left:=Left, Right:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered): Exit Sub
        
        Depthlimit = Depthlimit - 1
        p = PartitionWithItems_Delegate(Keys:=Keys, Items:=Items, Left:=Left, Right:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
        Call IntroSortWithItems_Delegate(Keys:=Keys, Items:=Items, Left:=p + 1, Right:=Right, Depthlimit:=Depthlimit, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
        Right = p - 1
    
    Loop

End Sub

Private Function PartitionWithItems_Delegate( _
                                                ByRef Keys() As Object, _
                                                ByVal Items As GenericArray, _
                                                ByVal Left As Long, _
                                                ByVal Right As Long, _
                                                ByVal Direction As SortOrder, _
                                                ByVal AscendingOrdered As IGenericPredicator) As Long
                                                
    Dim i As Long, j As Long, p As Long
    Dim Pivot As Object, tmp As Object
    
    p = PivotSwapWithItems_Delegate(Keys:=Keys, Items:=Items, i:=Left, j:=Right, Direction:=Direction, AscendingOrdered:=AscendingOrdered)
    Set Pivot = Keys(p) 'try to cast
   
    i = Left: j = Right - 1
    
    Set Keys(p) = Keys(j)
    Set Keys(j) = Pivot
    
    With Items
        
        Set tmp = .ElementAt(p)
        Set .ElementAt(p) = .ElementAt(j)
        Set .ElementAt(j) = tmp
    
        If (Direction = Ascending) Then
        
            Do While (i < j)
            
                Do: i = i + 1: Loop While (AscendingOrdered.IsTrue(Keys(i), Pivot))
                Do: j = j - 1: Loop While (AscendingOrdered.IsTrue(Pivot, Keys(j)))
                
                If (i >= j) Then _
                    Exit Do
            
                Set tmp = Keys(i)
                Set Keys(i) = Keys(j)
                Set Keys(j) = tmp
                
                Set tmp = .ElementAt(i)
                Set .ElementAt(i) = .ElementAt(j)
                Set .ElementAt(j) = tmp
                 
            Loop
            
        Else
        
            Do While (i < j)
            
                Do: i = i + 1: Loop While (AscendingOrdered.IsTrue(Pivot, Keys(i)))
                Do: j = j - 1: Loop While (AscendingOrdered.IsTrue(Keys(j), Pivot))
                
                If (i >= j) Then _
                    Exit Do
            
                Set tmp = Keys(i)
                Set Keys(i) = Keys(j)
                Set Keys(j) = tmp
                
                Set tmp = .ElementAt(i)
                Set .ElementAt(i) = .ElementAt(j)
                Set .ElementAt(j) = tmp
                 
            Loop
        
        End If
  
        j = (Right - 1)
        
        If Not (i = j) Then
        
            Set tmp = Keys(i)
            Set Keys(i) = Keys(j)
            Set Keys(j) = tmp
            
            Set tmp = Items.ElementAt(i)
            Set .ElementAt(i) = .ElementAt(j)
            Set .ElementAt(j) = tmp
            
        End If
        
    End With
 
    PartitionWithItems_Delegate = i
        
End Function
                              
Private Sub BinarySort( _
                         ByRef Keys() As Object, _
                         ByVal Left As Long, _
                         ByVal Right As Long, _
                         ByVal Direction As SortOrder)
    
    Dim Pivot As IGenericValue, tmp As Object
    Dim i As Long, j As Long, lo As Long, hi As Long, mid As Long
    
    If (Direction = Ascending) Then
    
        For i = Left + 1 To Right
    
            Set Pivot = Keys(i)
            If Not Pivot Is Nothing Then
    
                lo = Left
                hi = i
                
                With Pivot
                    Do
            
                        mid = lo + ((hi - lo) \ 2) '! truncate the floats
                        
                        Set tmp = Keys(mid)
                        If Not (tmp Is Nothing) Then
                            If .IsLessThan(tmp) Then
                                hi = mid
                            Else
                                lo = mid + 1
                            End If
                        Else
                            hi = mid
                        End If
                        
                    Loop While (lo < hi)
                End With
                
                For j = i - 1 To lo Step -1: Set Keys(j + 1) = Keys(j): Next
                Set Keys(lo) = Pivot
                
'                If (lo < i) Then
'
'                    j = i - 1
'                    Do
'                        Set Keys(j + 1) = Keys(j)
'                        j = j - 1
'                    Loop While (j >= lo)
'
'                    Set Keys(lo) = Pivot
'
'                End If

            End If
        Next
    
    Else
    
        For i = Left + 1 To Right
    
            Set Pivot = Keys(i)
            
            If Not Pivot Is Nothing Then
            
                lo = Left
                hi = i
                
                With Pivot
                    Do
            
                        mid = lo + ((hi - lo) \ 2) '! truncate the floats
                        
                        Set tmp = Keys(mid)
                        If Not (tmp Is Nothing) Then
                            If .IsGreaterThan(Keys(mid)) Then
                                hi = mid
                            Else
                                lo = mid + 1
                            End If
                        Else
                            hi = mid
                        End If
                        
                    Loop While (lo < hi)
                End With
            
                For j = i - 1 To lo Step -1: Set Keys(j + 1) = Keys(j): Next
                Set Keys(lo) = Pivot
                
            End If
        Next

    End If
 
End Sub

Private Sub HeapSort( _
                            ByRef Keys() As Object, _
                            ByVal lo As Long, _
                            ByVal hi As Long, _
                            ByVal Direction As SortOrder)

    Dim n As Long: n = hi - lo + 1
    Dim i As Long
    Dim tmp As Object
    
    If (Direction = Ascending) Then
        
        For i = n \ 2 To 1 Step -1: Call MaxHeap(Keys, i, n, lo): Next
        
        For i = n - 1 To 2 Step -1
            Set tmp = Keys(lo)
            Set Keys(lo) = Keys(lo + i)
            Set Keys(lo + i) = tmp
            
            Call MaxHeap(Keys, 1, i, lo)
        Next
        
    Else
    
        For i = n \ 2 To 1 Step -1: Call MinHeap(Keys, i, n, lo): Next
        
        For i = n - 1 To 2 Step -1
            Set tmp = Keys(lo)
            Set Keys(lo) = Keys(lo + i)
            Set Keys(lo + i) = tmp
            
            Call MinHeap(Keys, 1, i, lo)
        Next
        
    End If
End Sub

Private Sub MaxHeap(ByRef Keys() As Object, ByVal i As Long, ByVal n As Long, ByVal lo As Long)

    Dim Pivot As IGenericValue
    Dim Element As IGenericValue, tmp  As Object
    Dim Child As Long, Bound As Long
    
    Set Pivot = Keys(lo + i - 1) 'try to cast
    Bound = n \ 2 '! truncate the floats
    
    If Not (Pivot Is Nothing) Then
        
        With Pivot
            Do While (i <= Bound)
                Child = 2 * i
                Set Element = Keys(lo + Child - 1)
                If (Child < n) Then
                    Set tmp = Keys(lo + Child)
                    If Not (Element Is Nothing) Then
                        If Not (tmp Is Nothing) Then
                        
                            If Element.IsLessThan(tmp) Then _
                                Set Element = tmp: Child = Child + 1
                                
                        Else

                            Set Element = tmp: Child = Child + 1

                        End If
                    End If
                    
                End If

                If Not (Element Is Nothing) Then _
                    If .IsGreaterThan(Element) Then _
                        Exit Do
        
                Set Keys(lo + i - 1) = Element
                i = Child
            Loop
        End With
        
    Else
    
        Do While (i <= Bound)
            Child = 2 * i
            Set Element = Keys(lo + Child - 1)
            If (Child < n) Then
                Set tmp = Keys(lo + Child)
                
                If Not (Element Is Nothing) Then
                    If Not (tmp Is Nothing) Then
                    
                        If Element.IsLessThan(tmp) Then _
                            Set Element = tmp: Child = Child + 1

                    End If
                End If
                
            End If
           
            If Not (Element Is Nothing) Then _
                Exit Do

            Set Keys(lo + i - 1) = Element
            i = Child
        Loop
    
    End If

    Set Keys(lo + i - 1) = Pivot

End Sub

Private Sub MinHeap(ByRef Keys() As Object, ByVal i As Long, ByVal n As Long, ByVal lo As Long)

    Dim Pivot As IGenericValue
    Dim Element As IGenericValue, tmp  As Object
    Dim Child As Long, Bound As Long
    
    Set Pivot = Keys(lo + i - 1) 'try to cast
    Bound = n \ 2 '! truncate the floats
    
    If Not (Pivot Is Nothing) Then
        
        With Pivot
            Do While (i <= Bound)
                Child = 2 * i
                Set Element = Keys(lo + Child - 1)
                If (Child < n) Then
                    Set tmp = Keys(lo + Child)
                    If Not (Element Is Nothing) Then
                        If Not (tmp Is Nothing) Then
                        
                            If Element.IsGreaterThan(tmp) Then _
                                Set Element = tmp: Child = Child + 1
                                
                        Else

                            Set Element = tmp: Child = Child + 1

                        End If
                    End If
                    
                End If

                If Not (Element Is Nothing) Then _
                    If .IsLessThan(Element) Then _
                        Exit Do
        
                Set Keys(lo + i - 1) = Element
                i = Child
            Loop
        End With
        
    Else
    
        Do While (i <= Bound)
            Child = 2 * i
            Set Element = Keys(lo + Child - 1)
            If (Child < n) Then
                Set tmp = Keys(lo + Child)
                
                If Not (Element Is Nothing) Then
                    If Not (tmp Is Nothing) Then
                    
                        If Element.IsGreaterThan(tmp) Then _
                            Set Element = tmp: Child = Child + 1

                    End If
                End If
                
            End If
           
            If Not (Element Is Nothing) Then _
                Exit Do

            Set Keys(lo + i - 1) = Element
            i = Child
        Loop
    
    End If

    Set Keys(lo + i - 1) = Pivot

End Sub

Private Sub BinarySort_Delegate( _
                                    ByRef Keys() As Object, _
                                    ByVal Left As Long, _
                                    ByVal Right As Long, _
                                    ByVal Direction As SortOrder, _
                                    ByVal AscendingOrdered As IGenericPredicator)

    Dim Pivot As Object
    Dim i As Long, j As Long, lo As Long, hi As Long, mid As Long
  
    With AscendingOrdered
        If (Direction = Ascending) Then
        
            For i = Left + 1 To Right
    
                Set Pivot = Keys(i)
    
                lo = Left
                hi = i
    
                Do
    
                    mid = lo + ((hi - lo) \ 2) '! truncate the floats
    
                    If .IsTrue(Pivot, Keys(mid)) Then
                        hi = mid
                    Else
                        lo = mid + 1
                    End If
    
                Loop While (lo < hi)
                
                For j = i - 1 To lo Step -1: Set Keys(j + 1) = Keys(j): Next
                Set Keys(lo) = Pivot
                
            Next
        
        Else
        
            For i = Left + 1 To Right
    
                Set Pivot = Keys(i)
    
                lo = Left
                hi = i
    
                Do
    
                    mid = lo + ((hi - lo) \ 2) '! truncate the floats
    
                    If .IsTrue(Keys(mid), Pivot) Then
                        hi = mid
                    Else
                        lo = mid + 1
                    End If
    
                Loop While (lo < hi)
                
                For j = i - 1 To lo Step -1: Set Keys(j + 1) = Keys(j): Next
                Set Keys(lo) = Pivot
                
            Next
        
        End If
    End With

End Sub

Private Sub HeapSort_Delegate( _
                                ByRef Keys() As Object, _
                                ByVal Left As Long, _
                                ByVal Right As Long, _
                                ByVal Direction As SortOrder, _
                                ByVal AscendingOrdered As IGenericPredicator)

    Dim n As Long: n = Right - Left + 1
    Dim i As Long
    Dim tmp As Object
    
    If (Direction = Ascending) Then
    
        For i = n \ 2 To 1 Step -1: Call MaxHeap_Delegate(Keys, i, n, Left, AscendingOrdered): Next
        
        For i = n - 1 To 2 Step -1
            Set tmp = Keys(Left)
            Set Keys(Left) = Keys(Left + i)
            Set Keys(Left + i) = tmp
            
            Call MaxHeap_Delegate(Keys, 1, i, Left, AscendingOrdered)
        Next
        
    Else
    
        For i = n \ 2 To 1 Step -1: Call MinHeap_Delegate(Keys, i, n, Left, AscendingOrdered): Next
        
        For i = n - 1 To 2 Step -1
            Set tmp = Keys(Left)
            Set Keys(Left) = Keys(Left + i)
            Set Keys(Left + i) = tmp
            
            Call MinHeap_Delegate(Keys, 1, i, Left, AscendingOrdered)
        Next
        
    End If
    
End Sub

Private Sub MaxHeap_Delegate(ByRef Keys() As Object, ByVal i As Long, ByVal n As Long, ByVal lo As Long, ByVal AscendingOrdered As IGenericPredicator)

    Dim d As Object
    Dim Element As Object
    Dim Child As Long, Bound As Long
    
    Set d = Keys(lo + i - 1)
    Bound = n \ 2 '! truncate the floats
    
    With AscendingOrdered
        Do While (i <= Bound)
            Child = 2 * i
            Set Element = Keys(lo + Child - 1)
            If (Child < n) Then _
                If .IsTrue(Element, Keys(lo + Child)) Then _
                    Set Element = Keys(lo + Child): Child = Child + 1
        
            If .IsTrue(Element, d) Then _
                Exit Do
    
            Set Keys(lo + i - 1) = Element
            i = Child
        Loop
    End With
    
    Set Keys(lo + i - 1) = d

End Sub

Private Sub MinHeap_Delegate(ByRef Keys() As Object, ByVal i As Long, ByVal n As Long, ByVal lo As Long, ByVal AscendingOrdered As IGenericPredicator)

    Dim d As Object
    Dim Element As Object
    Dim Child As Long, Bound As Long
    
    Set d = Keys(lo + i - 1)
    Bound = n \ 2 '! truncate the floats
    
    With AscendingOrdered
        Do While (i <= Bound)
            Child = 2 * i
            Set Element = Keys(lo + Child - 1)
                      
            If (Child < n) Then _
                If .IsTrue(Keys(lo + Child), Element) Then _
                    Set Element = Keys(lo + Child): Child = Child + 1
            
            If .IsTrue(d, Element) Then _
                Exit Do
          
            Set Keys(lo + i - 1) = Element
            i = Child
        Loop
    End With
    
    Set Keys(lo + i - 1) = d

End Sub
                      
Private Sub BinarySortWithItems( _
                                    ByRef Keys() As Object, _
                                    ByVal Items As GenericArray, _
                                    ByVal Left As Long, _
                                    ByVal Right As Long, _
                                    ByVal Direction As SortOrder)

    Dim Pivot As IGenericValue, tmp As Object, Item As Object
    Dim i As Long, j As Long, lo As Long, hi As Long, mid As Long
    
    With Items
        If (Direction = Ascending) Then
        
            For i = Left + 1 To Right
        
                Set Pivot = Keys(i)
                Set Item = .ElementAt(i)
                
                If Not Pivot Is Nothing Then
        
                    lo = Left
                    hi = i
                    
                    With Pivot
                        Do
                
                            mid = lo + ((hi - lo) \ 2) '! truncate the floats
                            
                            Set tmp = Keys(mid)
                            If Not (tmp Is Nothing) Then
                                If .IsLessThan(tmp) Then
                                    hi = mid
                                Else
                                    lo = mid + 1
                                End If
                            Else
                                hi = mid
                            End If
                            
                        Loop While (lo < hi)
                    End With
                    
                    For j = i - 1 To lo Step -1
                        Set Keys(j + 1) = Keys(j)
                        Set .ElementAt(j + 1) = .ElementAt(j)
                    Next
                    
                    Set Keys(lo) = Pivot
                    Set .ElementAt(lo) = Item
                
                End If
            Next
    
        Else
        
            For i = Left + 1 To Right
        
                Set Pivot = Keys(i)
                Set Item = .ElementAt(i)
                
                If Not Pivot Is Nothing Then
                
                    lo = Left
                    hi = i
                    
                    With Pivot
                        Do
                
                            mid = lo + ((hi - lo) \ 2) '! truncate the floats
                            
                            Set tmp = Keys(mid)
                            If Not (tmp Is Nothing) Then
                                If .IsGreaterThan(Keys(mid)) Then
                                    hi = mid
                                Else
                                    lo = mid + 1
                                End If
                            Else
                                hi = mid
                            End If
                            
                        Loop While (lo < hi)
                    End With
                
                    For j = i - 1 To lo Step -1
                        Set Keys(j + 1) = Keys(j)
                        Set .ElementAt(j + 1) = .ElementAt(j)
                    Next
                    
                    Set Keys(lo) = Pivot
                    Set .ElementAt(lo) = Item
                
                End If
            Next
            
         End If
    End With
    
End Sub

Private Sub HeapSortWithItems( _
                                ByRef Keys() As Object, _
                                ByVal Items As GenericArray, _
                                ByVal lo As Long, _
                                ByVal hi As Long, _
                                ByVal Direction As SortOrder)

    Dim n As Long: n = hi - lo + 1
    Dim i As Long
    Dim tmp As Object
    
    With Items
        If Direction = Ascending Then
        
            For i = n \ 2 To 1 Step -1: Call MaxHeapWithItems(Keys, Items, i, n, lo): Next
            
            For i = n - 1 To 2 Step -1
                Set tmp = Keys(lo)
                Set Keys(lo) = Keys(lo + i)
                Set Keys(lo + i) = tmp
                
                Set tmp = .ElementAt(lo)
                Set .ElementAt(lo) = .ElementAt(lo + i)
                Set .ElementAt(lo + i) = tmp
                
                Call MaxHeapWithItems(Keys, Items, 1, i, lo)
            Next
            
        Else
        
            For i = n \ 2 To 1 Step -1: Call MinHeapWithItems(Keys, Items, i, n, lo): Next
            
            For i = n - 1 To 2 Step -1
                Set tmp = Keys(lo)
                Set Keys(lo) = Keys(lo + i)
                Set Keys(lo + i) = tmp
                
                Set tmp = .ElementAt(lo)
                Set .ElementAt(lo) = .ElementAt(lo + i)
                Set .ElementAt(lo + i) = tmp
                
                Call MinHeapWithItems(Keys, Items, 1, i, lo)
            Next
            
        End If
    End With
    
End Sub

Private Sub MaxHeapWithItems( _
                                ByRef Keys() As Object, _
                                ByRef Items As GenericArray, _
                                ByVal i As Long, _
                                ByVal n As Long, _
                                ByVal lo As Long)
    
    Dim Pivot As IGenericValue, Item As Object
    Dim Element As IGenericValue, tmp  As Object
    Dim Child As Long, Bound As Long
    
    Set Pivot = Keys(lo + i - 1) 'try to cast
    Bound = n \ 2 '! truncate the floats
    
    With Items
    
        Set Item = .ElementAt(lo + i - 1)
        
        If Not (Pivot Is Nothing) Then
            
            Do While (i <= Bound)
                Child = 2 * i
                Set Element = Keys(lo + Child - 1)
                If (Child < n) Then
                    Set tmp = Keys(lo + Child)
                    If Not (Element Is Nothing) Then
                        If Not (tmp Is Nothing) Then
                        
                            If Element.IsLessThan(tmp) Then _
                                Set Element = tmp: Child = Child + 1
                                
                        Else
    
                            Set Element = tmp: Child = Child + 1
    
                        End If
                    End If
                    
                End If
    
                If Not (Element Is Nothing) Then _
                    If Pivot.IsGreaterThan(Element) Then _
                        Exit Do
        
                Set Keys(lo + i - 1) = Element
                Set .ElementAt(lo + i - 1) = .ElementAt(lo + Child - 1)
                i = Child
            Loop
            
        Else
        
            Do While (i <= Bound)
                Child = 2 * i
                Set Element = Keys(lo + Child - 1)
                If (Child < n) Then
                    Set tmp = Keys(lo + Child)
                    
                    If Not (Element Is Nothing) Then
                        If Not (tmp Is Nothing) Then
                        
                            If Element.IsLessThan(tmp) Then _
                                Set Element = tmp: Child = Child + 1
    
                        End If
                    End If
                    
                End If
                
                If Not (Element Is Nothing) Then _
                    Exit Do
           
                Set Keys(lo + i - 1) = Element
                Set .ElementAt(lo + i - 1) = .ElementAt(lo + Child - 1)
                i = Child
            Loop
        
        End If

        Set Keys(lo + i - 1) = Pivot
        Set .ElementAt(lo + i - 1) = Item
        
    End With
   
End Sub

Private Sub MinHeapWithItems( _
                                ByRef Keys() As Object, _
                                ByRef Items As GenericArray, _
                                ByVal i As Long, _
                                ByVal n As Long, _
                                ByVal lo As Long)

    Dim Pivot As IGenericValue, Item As Object
    Dim Element As IGenericValue, tmp  As Object
    Dim Child As Long, Bound As Long
    
    Set Pivot = Keys(lo + i - 1) 'try to cast
    Bound = n \ 2 '! truncate the floats
    
    With Items
    
        Set Item = .ElementAt(lo + i - 1)
        
        If Not (Pivot Is Nothing) Then
            
            Do While (i <= Bound)
                Child = 2 * i
                Set Element = Keys(lo + Child - 1)
                If (Child < n) Then
                    Set tmp = Keys(lo + Child)
                    If Not (Element Is Nothing) Then
                        If Not (tmp Is Nothing) Then
                        
                            If Element.IsGreaterThan(tmp) Then _
                                Set Element = tmp: Child = Child + 1
                                
                        Else
    
                            Set Element = tmp: Child = Child + 1
    
                        End If
                    End If
                    
                End If
    
                If Not (Element Is Nothing) Then _
                    If Pivot.IsLessThan(Element) Then _
                        Exit Do
        
                Set Keys(lo + i - 1) = Element
                Set .ElementAt(lo + i - 1) = .ElementAt(lo + Child - 1)
                i = Child
            Loop
            
        Else
        
            Do While (i <= Bound)
                Child = 2 * i
                Set Element = Keys(lo + Child - 1)
                If (Child < n) Then
                    Set tmp = Keys(lo + Child)
                    
                    If Not (Element Is Nothing) Then
                        If Not (tmp Is Nothing) Then
                        
                            If Element.IsLessThan(tmp) Then _
                                Set Element = tmp: Child = Child + 1
    
                        End If
                    End If
                    
                End If
                
                If Not (Element Is Nothing) Then _
                    Exit Do
           
                Set Keys(lo + i - 1) = Element
                Set .ElementAt(lo + i - 1) = .ElementAt(lo + Child - 1)
                i = Child
            Loop
        
        End If

        Set Keys(lo + i - 1) = Pivot
        Set .ElementAt(lo + i - 1) = Item
        
    End With
    
End Sub

Private Sub BinarySortWithItems_Delegate( _
                                            ByRef Keys() As Object, _
                                            ByVal Items As GenericArray, _
                                            ByVal Left As Long, _
                                            ByVal Right As Long, _
                                            ByVal Direction As SortOrder, _
                                            ByVal AscendingOrdered As IGenericPredicator)

    Dim Pivot As Object, Item As Object
    Dim i As Long, j As Long, lo As Long, hi As Long, mid As Long
  
    With Items
        If (Direction = Ascending) Then
        
            For i = Left + 1 To Right
    
                Set Pivot = Keys(i)
                Set Item = .ElementAt(i)
                
                lo = Left
                hi = i
    
                Do
    
                    mid = lo + ((hi - lo) \ 2) '! truncate the floats
    
                    If AscendingOrdered.IsTrue(Pivot, Keys(mid)) Then
                        hi = mid
                    Else
                        lo = mid + 1
                    End If
    
                Loop While (lo < hi)
                
                For j = i - 1 To lo Step -1
                    Set Keys(j + 1) = Keys(j)
                    Set .ElementAt(j + 1) = .ElementAt(j)
                Next
                
                Set Keys(lo) = Pivot
                Set .ElementAt(lo) = Item
                
            Next
        
        Else
        
            For i = Left + 1 To Right
    
                Set Pivot = Keys(i)
                Set Item = .ElementAt(i)
                
                lo = Left
                hi = i
    
                Do
    
                    mid = lo + ((hi - lo) \ 2) '! truncate the floats
    
                    If AscendingOrdered.IsTrue(Keys(mid), Pivot) Then
                        hi = mid
                    Else
                        lo = mid + 1
                    End If
    
                Loop While (lo < hi)
                
                For j = i - 1 To lo Step -1
                    Set Keys(j + 1) = Keys(j)
                    Set .ElementAt(j + 1) = .ElementAt(j)
                Next
                
                Set Keys(lo) = Pivot
                Set .ElementAt(lo) = Item

            Next
        
        End If
    End With
   
End Sub

Private Sub HeapSortWithItems_Delegate( _
                                            ByRef Keys() As Object, _
                                            ByVal Items As GenericArray, _
                                            ByVal Left As Long, _
                                            ByVal Right As Long, _
                                            ByVal Direction As SortOrder, _
                                            ByVal AscendingOrdered As IGenericPredicator)

    Dim n As Long: n = Right - Left + 1
    Dim i As Long
    Dim tmp As Object
    
    With Items
        If (Direction = Ascending) Then
        
            For i = n \ 2 To 1 Step -1: Call MaxHeapWithItems_Delegate(Keys, Items, i, n, Left, AscendingOrdered): Next
            
            For i = n - 1 To 2 Step -1
                Set tmp = Keys(Left)
                Set Keys(Left) = Keys(Left + i)
                Set Keys(Left + i) = tmp
                
                Set tmp = .ElementAt(Left)
                Set .ElementAt(Left) = .ElementAt(Left + i)
                Set .ElementAt(Left + i) = tmp
                
                Call MaxHeapWithItems_Delegate(Keys, Items, 1, i, Left, AscendingOrdered)
            Next
            
        Else
        
            For i = n \ 2 To 1 Step -1: Call MinHeapWithItems_Delegate(Keys, Items, i, n, Left, AscendingOrdered): Next
            
            For i = n - 1 To 2 Step -1
                Set tmp = Keys(Left)
                Set Keys(Left) = Keys(Left + i)
                Set Keys(Left + i) = tmp
                
                Set tmp = .ElementAt(Left)
                Set .ElementAt(Left) = .ElementAt(Left + i)
                Set .ElementAt(Left + i) = tmp
                
                Call MinHeapWithItems_Delegate(Keys, Items, 1, i, Left, AscendingOrdered)
            Next
            
        End If
    End With
    
End Sub

Private Sub MaxHeapWithItems_Delegate( _
                                        ByRef Keys() As Object, _
                                        ByRef Items As GenericArray, _
                                        ByVal i As Long, _
                                        ByVal n As Long, _
                                        ByVal lo As Long, _
                                        ByVal AscendingOrdered As IGenericPredicator)

    Dim Pivot As Object
    Dim Item As Object
    Dim tmp As Object
    Dim Child As Long, Bound As Long
    
    Set Pivot = Keys(lo + i - 1)
    Bound = n \ 2 '! truncate the floats
    
    With Items
        
        Set Item = .ElementAt(lo + i - 1)
        
        Do While (i <= Bound)
            Child = 2 * i
            Set tmp = Keys(lo + Child - 1)
            If (Child < n) Then _
                If AscendingOrdered.IsTrue(tmp, Keys(lo + Child)) Then _
                    Set tmp = Keys(lo + Child): Child = Child + 1
        
            If AscendingOrdered.IsTrue(tmp, Pivot) Then _
                Exit Do
    
            Set Keys(lo + i - 1) = tmp
            Set .ElementAt(lo + i - 1) = .ElementAt(lo + Child - 1)
            i = Child
        Loop
   
        Set Keys(lo + i - 1) = Pivot
        Set .ElementAt(lo + i - 1) = Item
    
    End With
    
End Sub

Private Sub MinHeapWithItems_Delegate( _
                                        ByRef Keys() As Object, _
                                        ByRef Items As GenericArray, _
                                        ByVal i As Long, _
                                        ByVal n As Long, _
                                        ByVal lo As Long, _
                                        ByVal AscendingOrdered As IGenericPredicator)

    Dim Pivot As Object
    Dim Item As Object
    Dim tmp As Object
    Dim Child As Long, Bound As Long
    
    Set Pivot = Keys(lo + i - 1)
    Bound = n \ 2 '! truncate the floats

    With Items
    
        Set Item = .ElementAt(lo + i - 1)
        
        Do While (i <= Bound)
            Child = 2 * i
            Set tmp = Keys(lo + Child - 1)
                      
            If (Child < n) Then _
                If AscendingOrdered.IsTrue(Keys(lo + Child), tmp) Then _
                    Set tmp = Keys(lo + Child): Child = Child + 1
            
            If AscendingOrdered.IsTrue(Pivot, tmp) Then _
                Exit Do
          
            Set Keys(lo + i - 1) = tmp
            Set .ElementAt(lo + i - 1) = .ElementAt(lo + Child - 1)
            i = Child
        Loop
    
        Set Keys(lo + i - 1) = Pivot
        Set .ElementAt(lo + i - 1) = Item
        
    End With
    
End Sub

'Little trick t
Private Sub SwapInternal(ByRef A() As Object, ByVal i As Long, ByVal j As Long, Optional ByVal t As Object): Set t = A(i): Set A(i) = A(j): Set A(j) = t: End Sub

Private Function PivotSwap( _
                                ByRef Keys() As Object, _
                                ByVal i As Long, _
                                ByVal j As Long, _
                                ByVal Direction As SortOrder) As Long

    Dim mid As Long
    
    If Not i = j Then
    
        mid = i + ((j - i) \ 2)  '! truncate the floats
        
        Call InOrderSwap(Keys, i, mid, Direction)
        Call InOrderSwap(Keys, i, j, Direction)
        Call InOrderSwap(Keys, mid, j, Direction)
            
    End If
    
    PivotSwap = mid
    
End Function

Private Function InOrderSwap( _
                                ByRef Keys() As Object, _
                                ByVal i As Long, _
                                ByVal j As Long, _
                                ByVal Direction As SortOrder) As Boolean
    
    Dim Left As IGenericValue, Right As IGenericValue
    
    If (i = j) Then _
        Exit Function
    
     Set Left = Keys(i) 'Try cast
     Set Right = Keys(j) 'Try cast
     
     If Left Is Nothing Then
         If Not Right Is Nothing Then
             Set Keys(i) = Right
             Set Keys(j) = Left
             InOrderSwap = True
         End If
         Exit Function
     End If
    
     If Not (Right Is Nothing) Then
         If (Direction = Ascending) Then
             If Left.IsGreaterThan(Right) Then
                 Set Keys(i) = Right
                 Set Keys(j) = Left
                 InOrderSwap = True
             End If
         Else
             If Left.IsLessThan(Right) Then
                 Set Keys(i) = Right
                 Set Keys(j) = Left
                 InOrderSwap = True
             End If
         End If
     End If
        
End Function

Private Sub InOrderSwapWithItem( _
                                    ByRef Keys() As Object, _
                                    ByVal Items As GenericArray, _
                                    ByVal i As Long, _
                                    ByVal j As Long, _
                                    ByVal Direction As SortOrder)
    
    Dim tmp As Object
    
    If InOrderSwap(Keys, i, j, Direction) Then
        Set tmp = Items.ElementAt(i)
        Set Items.ElementAt(i) = Items.ElementAt(j)
        Set Items.ElementAt(j) = tmp
    End If
    
End Sub

Private Function PivotSwapWithItems( _
                                        ByRef Keys() As Object, _
                                        ByRef Items As GenericArray, _
                                        ByVal i As Long, _
                                        ByVal j As Long, _
                                        ByVal Direction As SortOrder) As Long
                                        
    Dim mid As Long
   
    If Not i = j Then
    
        mid = i + ((j - i) \ 2)  '! truncate the floats
            
        Call InOrderSwapWithItem(Keys, Items, i, mid, Direction)
        Call InOrderSwapWithItem(Keys, Items, i, j, Direction)
        Call InOrderSwapWithItem(Keys, Items, mid, j, Direction)
    
    End If
    
    PivotSwapWithItems = mid
    
End Function

Private Function InOrderSwap_Delegate( _
                                            ByRef Keys() As Object, _
                                            ByVal i As Long, _
                                            ByVal j As Long, _
                                            ByVal Direction As SortOrder, _
                                            ByVal AscendingOrdered As IGenericPredicator) As Boolean
    
    Dim tmp As Object
    
    If (i = j) Then _
        Exit Function
    
    Set tmp = Keys(i)
    
    If (Direction = Ascending) Then
        InOrderSwap_Delegate = AscendingOrdered.IsTrue(Keys(j), tmp)
    Else
        InOrderSwap_Delegate = AscendingOrdered.IsTrue(tmp, Keys(j))
    End If
    
    If InOrderSwap_Delegate Then
        Set Keys(i) = Keys(j)
        Set Keys(j) = tmp
    End If
    
End Function

Private Function PivotSwap_Delegate( _
                                        ByRef Keys() As Object, _
                                        ByVal i As Long, _
                                        ByVal j As Long, _
                                        ByVal Direction As SortOrder, _
                                        ByVal AscendingOrdered As IGenericPredicator) As Long
                                        
    Dim mid As Long
    
    If Not i = j Then
    
        mid = i + ((j - i) \ 2)  '! truncate the floats
        
        Call InOrderSwap_Delegate(Keys, i, mid, Direction, AscendingOrdered)
        Call InOrderSwap_Delegate(Keys, i, j, Direction, AscendingOrdered)
        Call InOrderSwap_Delegate(Keys, mid, j, Direction, AscendingOrdered)
            
    End If
    
    PivotSwap_Delegate = mid
    
End Function

 Private Function PivotSwapWithItems_Delegate( _
                                                ByRef Keys() As Object, _
                                                ByRef Items As GenericArray, _
                                                ByVal i As Long, _
                                                ByVal j As Long, _
                                                ByVal Direction As SortOrder, _
                                                ByVal AscendingOrdered As IGenericPredicator) As Long
    
    Dim mid As Long
    
    If Not i = j Then
    
        mid = i + ((j - i) \ 2)  '! truncate the floats
        
        Call InOrderSwapWithItem_Delegate(Keys, Items, i, mid, Direction, AscendingOrdered)
        Call InOrderSwapWithItem_Delegate(Keys, Items, i, j, Direction, AscendingOrdered)
        Call InOrderSwapWithItem_Delegate(Keys, Items, mid, j, Direction, AscendingOrdered)
            
    End If
    
    PivotSwapWithItems_Delegate = mid
    
End Function

Private Sub InOrderSwapWithItem_Delegate( _
                                            ByRef Keys() As Object, _
                                            ByVal Items As GenericArray, _
                                            ByVal i As Long, _
                                            ByVal j As Long, _
                                            ByVal Direction As SortOrder, _
                                            ByVal AscendingOrdered As IGenericPredicator)
    
    Dim tmp As Object
    
    If InOrderSwap_Delegate(Keys, i, j, Direction, AscendingOrdered) Then
        Set tmp = Items.ElementAt(i)
        Set Items.ElementAt(i) = Items.ElementAt(j)
        Set Items.ElementAt(j) = tmp
    End If
    
End Sub

Private Sub SwapWithItemInternal(ByRef A() As Object, ByRef Items As GenericArray, ByVal i As Long, ByVal j As Long)

    Dim tmp As Object

    Set tmp = A(i): Set A(i) = A(j): Set A(j) = tmp
    
    Set tmp = Items.ElementAt(i)
    Set Items.ElementAt(i) = Items.ElementAt(j)
    Set Items.ElementAt(j) = tmp

End Sub

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericArray: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = GenericArray: End Property

Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(This): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(This): End Property
Private Property Get IGeneric_HashCode() As Long:  IGeneric_HashCode = ObjPtr(Me): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = True: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GenericArray):  End Property

'@Ignore EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As Object) As Boolean:: IGeneric_Equals = Me Is Other: End Function

Private Function IGeneric_ToString() As String
    IGeneric_ToString = _
                        "Class: " & VBA.TypeName$(Me) & vbNewLine & _
                        "Rank: " & This.SafeArray.cDims & vbNewLine & _
                        "Elements: " & This.Length
End Function

'IGenericCollection
Private Property Get IGenericCollection_HasRandomAccess() As Boolean: IGenericCollection_HasRandomAccess = True: End Property
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = This.Length: End Property

Private Function IGenericCollection_GetAt(ByVal Index As Long) As Object
    
    Dim x As Span
    
    If (Index >= This.Length) Then _
        Call GenericError.Build(Me, "Private Function IGenericCollection_GetAt").IndexOutofRange("(Index >= This.Length)")
    
    If (This.SafeArray.cDims = 1) Then
        
        Set IGenericCollection_GetAt = This.Internal(Index)
    
    Else

        x.pvData = This.SafeArray.pvData + (Index * This.SafeArray.cbElements)
        x.fFeatures = This.SafeArray.fFeatures
        x.cbElements = This.SafeArray.cbElements
        x.cDims = 1
        x.cElements = This.Length

        If (SafeArrayGetElement(ByVal VarPtr(x), 0&, IGenericCollection_GetAt) = DISP_E_BADINDEX) Then _
            Call GenericError.Build(Me, "Private Function IGenericCollection_GetAt").InvalidInput("Invalid Index")

    End If

End Function

Private Sub IGenericCollection_Clear()
  
    If (This.Length = 0) Then _
        Exit Sub
        
    If (This.Descriptor = 0) Then _
        Call GenericError.Build(Me, "Private Sub IGenericCollection_Clear").InvalidOperation("(This.Descriptor = 0)")
  
    If Not (SafeArrayDestroyData(ByVal This.Descriptor)) = S_OK Then _
        Call GenericError.Build(Me, "Private Sub IGenericCollection_Clear").InvalidOperation("SafeArrayDestroyData") 'if not fFeatures & FADF_STATIC then pvData-> null

    This.SafeArray.pvData = 0
    
    If Not (SafeArrayAllocData(ByVal This.Descriptor)) = S_OK Then _
        Call GenericError.Build(Me, "Private Sub IGenericCollection_Clear").InvalidOperation("SafeArrayAllocData")
    
    If Not (SafeArrayAccessData(ByVal This.Descriptor, This.SafeArray.pvData)) = S_OK Then _
        Call GenericError.Build(Me, "Private Sub IGenericCollection_Clear").InvalidOperation("SafeArrayAccessData")
        
    If Not (SafeArrayUnaccessData(ByVal This.Descriptor)) = S_OK Then _
        Call GenericError.Build(Me, "Private Sub IGenericCollection_Clear").InvalidOperation("SafeArrayUnaccessData")

End Sub

Private Function IGenericCollection_ToArray() As GenericArray: Set IGenericCollection_ToArray = Me.Elements.Copy: End Function

Private Function IGenericCollection_Contains(ByVal Element As Object) As Boolean
    
    '@Ignore VariableNotAssigned
    Dim Candidate As Object
    Dim Casted As IGenericValue
    Dim i As Long, x As Span, Ptr As Long
    
    If (This.Length = 0) Then _
        Exit Function
        
    If (This.SafeArray.cDims = 1) Then
        
        If (IndexOfInternal(Objects:=This.Internal, Element:=Element, LowerBound:=MIN_INDEX, UpperBound:=This.Length + MIN_INDEX - 1, Predicate:=Nothing) = NOT_FOUND) Then _
            Exit Function
    
    Else
        
        x.pvData = This.SafeArray.pvData
        x.fFeatures = This.SafeArray.fFeatures
        x.cbElements = This.SafeArray.cbElements
        x.cDims = 1
        x.cElements = This.Length
    
        Ptr = VarPtr(x)
        
        If Not (Element Is Nothing) Then
            Set Casted = Element 'try cast
            With Casted
                Do
                    If (SafeArrayGetElement(ByVal Ptr, 0&, ByVal Candidate) = DISP_E_BADINDEX) Then _
                        Call GenericError.Build(Me, "Private Function IGenericCollection_Contains").InvalidInput("Invalid Index")
                    
                    If .Equals(Candidate) Then _
                        Exit Do
        
                    x.pvData = x.pvData + ELEMENT_SIZE
                    i = i + 1
                Loop While i < x.cElements
            End With
            
        Else
             
             Do
                If (SafeArrayGetElement(ByVal Ptr, 0&, ByVal Candidate) = DISP_E_BADINDEX) Then _
                    Call GenericError.Build(Me, "Private Function IGenericCollection_Contains").InvalidInput("Invalid Index")
                
                If (Candidate Is Nothing) Then _
                    Exit Do
    
                x.pvData = x.pvData + ELEMENT_SIZE
                i = i + 1
            Loop While i < x.cElements
        
        End If
        
        If (i = x.cElements) Then _
            Exit Function
                        
    End If
    
    IGenericCollection_Contains = True
    
End Function

Private Sub IGenericCollection_CopyTo(ByVal Target As GenericArray, ByVal Index As Long): Call Me.CopyTo(Target:=Target, Index:=Index, BeginAt:=MIN_INDEX, Count:=This.Length): End Sub

Private Function IGenericCollection_Copy() As IGenericCollection
    
    Dim Clone As Member
    
    Set IGenericCollection_Copy = GenericArray 'avoid null
    
    If (Me.IsEmpty) Then _
        Exit Function
  
    With Clone
    
        .Length = This.Length
        .SafeArray.cDims = This.SafeArray.cDims
        .SafeArray.fFeatures = This.SafeArray.fFeatures
        .SafeArray.cbElements = This.SafeArray.cbElements
        .SafeArray.Bounds = This.SafeArray.Bounds

        If Not (SafeArrayCopy(ByVal This.Descriptor, .Descriptor)) = S_OK Then _
            Call GenericError.Build(Me, "Private Function IGenericCollection_Copy").InvalidOperation("SafeArrayCopy")

        If Not (SafeArrayAccessData(ByVal .Descriptor, .SafeArray.pvData)) = S_OK Then _
            Call GenericError.Build(Me, "Private Function IGenericCollection_Copy").InvalidOperation("SafeArrayAccessData")
        
        If Not (SafeArrayUnaccessData(ByVal .Descriptor)) = S_OK Then _
            Call GenericError.Build(Me, "Private Function IGenericCollection_Copy").InvalidOperation("SafeArrayUnaccessData")
        
        'Everything is fine, load internal
        Call CopyMemory(Dst:=ByVal VarPtrArray(.Internal), Src:=ByVal VarPtr(.Descriptor), Length:=POINTERSIZE)
    
    End With
    
    Set IGenericCollection_Copy = New GenericArray
    Call System.Inject(IGenericCollection_Copy, VarPtr(Clone), LenB(Clone))
    
End Function

Private Function IGenericCollection_Iterator() As IGenericIterator: Set IGenericCollection_Iterator = Me.Range: End Function

'IGenericSequence_Iterator
Private Function IGenericSequence_Iterator() As IGenericIterator: Set IGenericSequence_Iterator = Me.Range: End Function


