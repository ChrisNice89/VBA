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
Option Compare Database

Option Explicit
'@PredeclaredId

Implements IPrototype
Implements IGeneric
Implements IGenericList
Implements IGenericCollection

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

Private Const OPTION_BASE As Byte = 1 'Option Base { 0 | 1 }
Private Const MAX_ARRAY_LENGHT As Long = &H7FEFFFFF

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long

Private Type SAFEARRAYBOUND
    cElements    As Long
    lLbound      As Long
End Type

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
#End If
 
Private Type Member
   Struct As SAFEARRAYSTRUCT
   Descriptor As LongPtr
   Size As LongPtr
   Elements As Long
   Indices() As Long
   ptrIndices As LongPtr
   Internal() As IGeneric
End Type
Private myClass As Member

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get MaxLength() As Long: MaxLength = MAX_ARRAY_LENGHT: End Property
Public Property Get LowerBound() As Long: LowerBound = OPTION_BASE: End Property
Public Property Get IsEmpty() As Boolean: IsEmpty = (myClass.Descriptor = 0): End Property
Public Property Get Descriptor() As LongPtr: Descriptor = myClass.Descriptor: End Property
'Public Property Get Count(Optional ByVal Dimension As Long = 1) As Long: Call Catch(SafeArrayGetUBound(ByVal myClass.Descriptor, Dimension, Count)): End Property 'at least 1!
Public Property Get LengthOf(ByVal Dimension As Long) As Long:: Call Catch(SafeArrayGetUBound(ByVal myClass.Descriptor, Dimension, LengthOf)): End Property
Public Property Get Length() As Long: Length = myClass.Elements: End Property
Public Property Get TotalSize() As LongPtr: TotalSize = myClass.Size: End Property
Public Property Get Rank() As Long: Rank = SafeArrayGetDim(ByVal myClass.Descriptor): End Property

Public Property Set ElementAt(ByVal Row As Long, Optional ByVal Column As Long = OPTION_BASE, ByVal Value As IGeneric)
    myClass.Indices(0) = Row: myClass.Indices(1) = Column

    If (SafeArrayPutElement(ByVal myClass.Descriptor, ByVal myClass.ptrIndices, ByVal ObjPtr(Value)) = DISP_E_BADINDEX) Then _
        Call Skynet.Push(Me, "Set Object").InvalidInput.Raise("Invalid Index")
    
'       Set myClass.Internal(Row, Column) = Value
End Property

Public Property Get ElementAt(ByVal Row As Long, Optional ByVal Column As Long = OPTION_BASE) As IGeneric
Attribute ElementAt.VB_UserMemId = 0
    myClass.Indices(0) = Row: myClass.Indices(1) = Column

    If (SafeArrayGetElement(ByVal myClass.Descriptor, ByVal myClass.ptrIndices, ElementAt) = DISP_E_BADINDEX) Then _
        Call Skynet.Push(Me, "Get Object").InvalidInput.Raise("Invalid Index")
        
'       Set ElementAt = myClass.Internal(Row, Column)
End Property

Public Function AdressOf(ByVal Element As Long) As LongPtr
    AdressOf = myClass.Struct.pvData + (Me.SizeOf(Element - OPTION_BASE))
End Function

Public Function SizeOf(ByVal Elements As Long) As Long
    SizeOf = Elements * myClass.Struct.cbElements
End Function

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

Private Sub SwitchTo(ByVal Descriptor As LongPtr, ByRef S As SAFEARRAYSTRUCT)
    With Skynet
        Call .CopyData(ByVal Descriptor, VarPtr(S), 16&): _
        Call .CopyData(ByVal Descriptor + 16&, VarPtr(S.Bounds(LBound(S.Bounds))), S.cDims * 8&)
    End With
End Sub

Public Function Build(ByVal Rows As Long, Optional ByVal Columns As Long = OPTION_BASE) As GenericArray
    
    Dim SAB() As SAFEARRAYBOUND
    Dim i As Long
    
    If (Rows < OPTION_BASE) Or (Columns < OPTION_BASE) Then _
        Call Skynet.Push(Me, "Build").ArgOutOfRange.Raise("A zero(one)-based array indexed array has its first item indexed as 0(1)")
       
    If Columns > 1 Then
       ReDim SAB(1 To 2)
       SAB(2).cElements = Columns
       SAB(2).lLbound = OPTION_BASE
    Else
        ReDim SAB(1 To 1)
    End If
    
    With SAB(1)
        .cElements = Rows
        .lLbound = OPTION_BASE
    End With
    
    With myClass
        .Descriptor = SafeArrayCreateEx(VbVarType.vbObject, UBound(SAB), SAB(1), ByVal 0&)
'        ReDim .Internal(1 To Rows, 1 To Columns)
'        .Descriptor = Skynet.PointerOf(.Internal)
        If (.Descriptor = 0) Then _
            Call Skynet.Push(Me, "Build").InvalidOperation.Raise("SafeArrayCreateEx returned zero")
        
        ReDim .Indices(1)
        .ptrIndices = VarPtr(.Indices(0))
        .Struct = GetArrayStructure(.Descriptor)
        .Elements = 1
        .Size = 1
        
        For i = LBound(.Struct.Bounds) To UBound(.Struct.Bounds)
            .Elements = .Elements * .Struct.Bounds(i).cElements
            .Size = .Size * .Struct.Bounds(i).cElements * .Struct.cbElements
        Next

    End With
    
    Set Build = Skynet.Create(Me)
       
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
End Sub

Private Sub Class_Terminate()
    If Not (Me.IsEmpty) Then
    
        Call Skynet.Generic(Me).Dispose
    End If
End Sub

Public Sub Clear()
    Call Catch(SafeArrayDestroyData(ByVal Me.Descriptor))
End Sub

Public Function SlizeColumn(ByVal Index As Long) As GenericArray
    
    Dim Overlay As SAFEARRAYSTRUCT
    Dim Column As GenericArray
    
    If (Me.Rank <> 2) Then _
        Call Skynet.Push(Me, "IndexOf").NotSupported.Raise("Me.Rank != 2")
            
    If (Index < 1) Or (Index > Me.LengthOf(2)) Then _
        Call Skynet.Push(Me, "SlizeColumn").IndexOutofRange.Raise("Index not valid")
    
    If (Me.IsEmpty) Then _
        Call Skynet.Push(Me, "SlizeColumn").NULLPOINTER.Raise("Me.IsEmpty")
        
    Set Column = GenericArray.Build(Me.LengthOf(1))
  
    Overlay = GetArrayStructure(Column.Descriptor)
    Overlay.pvData = Me.AdressOf(Element:=Me.LengthOf(1) * Index)
    
    With Skynet
        Call SwitchTo(Me.Descriptor, Overlay)
        Call Catch(SafeArrayCopyData(ByVal Me.Descriptor, ByVal Column.Descriptor))
        Call SwitchTo(Me.Descriptor, myClass.Struct)
    End With
    
    Set SlizeColumn = Column
    
End Function

Public Sub CopyTo( _
                    ByVal GArray As GenericArray, _
                    ByVal ArrayIndex As Long, _
                    ByVal SourceIndex As Long, _
                    Optional ByVal Elements As Long = 0)
                    
    Dim Overlay As SAFEARRAYSTRUCT
 
    If (Elements < 0) Then _
        Call Skynet.Push(Me, "CopyTo").InvalidInput.Raise("Elements to copy  <  0")

    If (Elements = 0) Then _
        Elements = Me.Length
    
'    If (SourceAdress < Me.DataPointer) Or (SourceAdress > Me.DataPointer + Me.Size) Then _
'        Call Skynet.Push(Me, "CopyTo").InvalidInput.Raise("Invalid SourceAdress")
'
'    If ((Me.DataPointer + Me.Size) Mod SourceAdress) / myClass.Struct.cbElements < Elements Then _
'        Call Skynet.Push(Me, "CopyTo").InvalidInput.Raise("Elements to copy > source start - source end")
'
'    If (GA Is Nothing) Or (GA.IsEmpty) Then _
'        Call Skynet.Push(Me, "CopyTo").NULLPOINTER.Raise("GA Is Nothing")
'
'    If (GAAdress < GA.DataPointer) Or (GAAdress > GA.DataPointer + GA.Size) Then _
'        Call Skynet.Push(Me, "CopyTo").InvalidInput.Raise("Invalid GAAdress")
'
'    If ((GA.DataPointer + GA.Size) Mod GAAdress) / myClass.Struct.cbElements < Elements Then _
'        Call Skynet.Push(Me, "CopyTo").InvalidInput.Raise("Elements to copy > GA start - GA end")
    
    With Overlay
        .pvData = Me.AdressOf(SourceIndex)
        .cDims = 1
        .cbElements = myClass.Struct.cbElements
        .fFeatures = myClass.Struct.fFeatures
        ReDim .Bounds(1 To 1)
        .Bounds(1).cElements = Elements
        .Bounds(1).lLbound = OPTION_BASE
    End With
    
    Call SwitchTo(myClass.Descriptor, Overlay)
   
    With GArray
        Overlay.pvData = .AdressOf(ArrayIndex)
        Call SwitchTo(.Descriptor, Overlay)
        Call Catch(SafeArrayCopyData(ByVal myClass.Descriptor, ByVal .Descriptor))
        Call .Refresh
    End With
    
    Call SwitchTo(myClass.Descriptor, myClass.Struct)
End Sub

Public Sub Refresh()
    Call SwitchTo(myClass.Descriptor, myClass.Struct)
End Sub

Public Function Test()
    Dim x() As IGeneric
    'ReDim x(1 To 10, 1 To 2)
    
    Set Me.ElementAt(1) = GString("Test")
    Call SwitchTo(Skynet.PointerOf(x), myClass.Struct)
    
    Debug.Print x(1).ToString
End Function

Private Function GetArrayStructure(ByVal SafeArrayPointer As LongPtr) As SAFEARRAYSTRUCT
    
    If (SafeArrayPointer = 0) Then _
        Skynet.Push(Me, "GetArrayStructure").NULLPOINTER.Raise ("SafeArrayPointer")

    Call Skynet.CopyData(VarPtr(GetArrayStructure), SafeArrayPointer, 16&)
    
    With GetArrayStructure
        If Not (.cDims > 0) Then _
            Call Skynet.Push(Me, "GetArrayStructure").Overflow.Raise("Not (.cDims > 0)")
        
        ReDim .Bounds(1 To .cDims)
        Call Skynet.CopyData(VarPtr(.Bounds(1)), ByVal (SafeArrayPointer + 16&), .cDims * 8&)
    End With

End Function

'Returns the index of the first occurrence of a given value in a range of
'this list. It does a linear, O(n) search. The list is searched forwards, starting at index
'and upto count number of elements. The
'elements of the list are compared to the given value using the
'Object.Equals method.

Public Function IndexOf( _
                            ByVal Value As IGeneric, _
                            Optional ByVal StartIndex As Long = OPTION_BASE, _
                            Optional ByVal Count As Long, _
                            Optional ByVal CheckArguments As Boolean = True) As Long
    
    Dim i As Long
    Dim EndIndex As Long
    Dim Candidate As IGeneric

    If CheckArguments Then
        If Not (Me.Rank = 1) Then _
            Call Skynet.Push(Me, "IndexOf").NotSupported.Raise("Me.Rank != 1")
            
        Dim lb As Long: lb = OPTION_BASE
        If (StartIndex < lb) Or (StartIndex > Me.LengthOf(1) + lb) Then _
            Call Skynet.Push(Me, "IndexOf").IndexOutofRange.Raise("(StartIndex < lb) Or (StartIndex > Me.Count() + lb)")
    
        If (Count = 0) Then
            Count = (Me.LengthOf(1) - StartIndex) + lb
        Else
            If (Count < 0) Or (Count > Me.LengthOf(1) - StartIndex + lb) Then _
                Call Skynet.Push(Me, "IndexOf").IndexOutofRange.Raise("(Count < 0) Or (Count > Me.elements - StartIndex + lb)")
        End If
    End If
    
    EndIndex = StartIndex + Count - OPTION_BASE
    
        
    If Not (Value Is Nothing) Then
        For i = StartIndex To EndIndex
            myClass.Indices(0) = i
            Call SafeArrayGetElement(ByVal myClass.Descriptor, ByVal myClass.ptrIndices, Candidate)
            If Not Candidate Is Nothing Then _
                If Candidate.Equals(Value) Then IndexOf = i: Exit Function
        Next
    Else
        For i = StartIndex To EndIndex
            myClass.Indices(0) = i
            Call SafeArrayGetElement(ByVal myClass.Descriptor, ByVal myClass.ptrIndices, Candidate)
            If Candidate Is Nothing Then _
                IndexOf = i: Exit Function
        Next
    End If
    
    IndexOf = -1
    
End Function

Public Sub Transpose()
    
    Const ZEROPOINTER As Long = 0
    Dim ptrT1 As LongPtr, ptrT2 As LongPtr
    Dim t1() As ObjectTransfer, t2() As ObjectTransfer
    Dim i As Long, j As Long, M As Long, N As Long
    
    If (Me.IsEmpty) Then _
        Call Skynet.Push(Me, "Transpose").NULLPOINTER.Raise("SafeArray not allocated")
        
    If (Me.Rank <> 2) Then _
        Call Skynet.Push(Me, "Transpose").InvalidOperation.Raise("Rank <> 2")
        
    With myClass.Struct
        M = .Bounds(2).cElements
        N = .Bounds(1).cElements
        
        ReDim t1(1 To M, 1 To N)
        ReDim t2(1 To N, 1 To M)

        ptrT1 = VarPtr(t1(1, 1))
        ptrT2 = VarPtr(t2(1, 1))

        Call Skynet.CopyData(Dst:=ByVal ptrT1, Src:=ByVal .pvData, Bytes:=Me.TotalSize)
        
        For i = LBound(t1, 1) To UBound(t1, 1)
            For j = LBound(t1, 2) To UBound(t1, 2)
                t2(j, i) = t1(i, j)
            Next
        Next
    
        Call Skynet.CopyData(Dst:=ByVal .pvData, Src:=ByVal ptrT2, Bytes:=Me.TotalSize)

        Dim Bounds() As SAFEARRAYBOUND
        ReDim Bounds(1 To 2)
        Bounds(1) = .Bounds(2): Bounds(2) = .Bounds(1)

        Call Skynet.CopyData(Dst:=ByVal Me.Descriptor + 16&, Src:=VarPtr(Bounds(1)), Bytes:=.cDims * 8&)
        .Bounds(1) = Bounds(2): .Bounds(2) = Bounds(1)
     
    End With

'    With Skynet
'        Call .CopyData(ByVal ptrT1, ByVal VarPtr(ZEROPOINTER), .PointerLength)
'        Call .CopyData(ByVal ptrT2, ByVal VarPtr(ZEROPOINTER), .PointerLength)
'    End With
'
End Sub

Public Sub Sort(Optional ByVal Column As Long = 1)

End Sub

Public Sub Feature()
    
    Const FADF_AUTO = &H1
    Const FADF_STATIC = &H2
    Const FADF_EMBEDDED = &H4
    Const FADF_FIXEDSIZE = &H10
    Const FADF_RECORD = &H20
    Const FADF_HAVEIID = &H40
    Const FADF_HAVEVARTYPE = &H80
    Const FADF_BSTR = &H100
    Const FADF_UNKNOWN = &H200
    Const FADF_DISPATCH = &H400
    Const FADF_VARIANT = &H800
    
    
    Debug.Print myClass.Struct.fFeatures

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

'IGeneric
Private Sub IGeneric_Dispose()
    If (SafeArrayUnlock(myClass.Descriptor) = S_OK) Then
        Call Catch(SafeArrayDestroyData(ByVal Me.Descriptor))
        Call Catch(SafeArrayDestroyDescriptor(ByVal Me.Descriptor))
'    Call Catch(SafeArrayDestroy(ByVal Descriptor))
    End If
    myClass = EmptyStructure
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member: c = myClass
    c.Descriptor = 0
'    With GetArrayStructure(Skynet.PointerOf(c.List))
'
'    End With
'    c.Descriptor = Skynet.PointerOf(c.List)
    
    Call Catch(SafeArrayCopy(ByVal myClass.Descriptor, c.Descriptor))
    
    Virtual.Pointer = VarPtr(c)
    Set IGeneric_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
    
End Function

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

'IGenericList
Private Property Get IGenericList_Elements() As IGenericCollection: Set IGenericList_Elements = Me: End Property
Private Property Get IGenericList_Item(ByVal Index As Long) As IGeneric: Set IGenericList_Item = Me(Index): End Property
Private Property Set IGenericList_Item(ByVal Index As Long, ByVal Item As IGeneric): Set Me(Index) = Item: End Property
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = IGeneric_HashValue = ObjPtr(Me): End Property
Private Function IGenericList_IndexOf(ByVal Item As IGeneric) As Long: IGenericList_IndexOf = Me.IndexOf(Value:=Item, StartIndex:=Me.LowerBound, Count:=Me.Length): End Function

'IGCollection
Private Property Get IGenericCollection_Count() As Long: IGenericCollection_Count = Me.Length: End Property
Private Property Get IGenericCollection_IsFixedSize() As Boolean: IGenericCollection_IsFixedSize = True: End Property
Private Property Get IGenericCollection_IsReadOnly() As Boolean: IGenericCollection_IsReadOnly = False: End Property
Private Sub IGenericCollection_Clear(): Call Me.Clear: End Sub
Private Sub IGenericCollection_CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): Call Me.CopyTo(GArray:=GArray, ArrayIndex:=1, SourceIndex:=Me.LowerBound, Elements:=Me.Length): End Sub

Private Function IGenericCollection_NewEnumerator() As IGenericEnumerator

End Function

'IPrototype
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GenericArray: End Function
Private Property Get IPrototype_Default() As Object: Set IPrototype_Default = GenericArray: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property


