VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "<T>Class"
Option Compare Text

Implements IGeneric
Implements IGenericComparer

Public Enum ISqlDataType
    AUTOINCREMENT = -1
    Empty_Type = 0 'No value
    SmallInt_Type = 2 ' A 2-byte signed integer.
    Integer_Type = 3 'A 4-byte signed integer.
    Single_Type = 4 'A single-precision floating-point value.
    Double_Type = 5 'A double-precision floating-point value.
    Currency_Type = 6 'A currency value
    Date_Type = 7 'The number of days since December 30, 1899 + the fraction of a day.
    BSTR_Type = 8 'A null-terminated character string.
    IDispatch_Type = 9 'A pointer to an IDispatch interface on a COM object. Note: Currently not supported by ADO.
    Error_Type = 10 'A 32-bit error code
    Boolean_Type = 11 'A boolean value.
    Variant_Type = 12 'An Automation Variant. Note: Currently not supported by ADO.
    IUnknown_Type = 13 'A pointer to an IUnknown interface on a COM object. Note: Currently not supported by ADO.
    Decimal_Type = 14 'An exact numeric value with a fixed precision and scale.
    TinyInt_Type = 16 'A 1-byte signed integer.
    UnsignedTinyInt_Type = 17 'A 1-byte unsigned integer.
    UnsignedSmallInt_Type = 18 'A 2-byte unsigned integer.
    UnsignedInt_Type = 19 'A 4-byte unsigned integer.
    BigInt_Type = 20 'An 8-byte signed integer.
    UnsignedBigInt_Type = 21 'An 8-byte unsigned integer.
    FileTime_Type = 64 'The number of 100-nanosecond intervals since January 1,1601
    GUID_Type = 72 'A globally unique identifier (GUID)
    Binary_Type = 128 'A binary value.
    Char_Type = 129 'A string value.
    WChar_Type = 130 'A null-terminated Unicode character string.
    Numeric_Type = 131 'An exact numeric value with a fixed precision and scale.
    UserDefined_Type = 132 'A user-defined variable.
    DBDate_Type = 133 'A date value (yyyymmdd).
    DBTime_Type = 134 'A time value (hhmmss).
    DBTimeStamp_Type = 135 'A date/time stamp (yyyymmddhhmmss plus a fraction in billionths).
    Chapter_Type = 136 'A 4-byte chapter value that identifies rows in a child rowset
    PropVariant_Type = 138 'An Automation PROPVARIANT.
    VarNumeric_Type = 139 'A numeric value (Parameter object only).
    VarChar_Type = 200 'A string value (Parameter object only).
    LongVarChar_Type = 201 'A long string value.
    VarWChar_Type = 202 'A null-terminated Unicode character string.
    LongVarWChar_Type = 203 'A long null-terminated Unicode string value.
    VarBinary_Type = 204 'A binary value (Parameter object only).
    LongVarBinary_Type = 205 'A long binary value.
    'AdArray = 0x2000 'A flag value combined with another data
End Enum

Public Property Get ToValue() As Variant: ToValue = Null: End Property
Attribute ToValue.VB_UserMemId = 0
Public Property Get SqlType() As ISqlDataType: SqlType = IUnknown_Type: End Property
Public Property Get VarType() As VbVarType: End Property
Public Property Get Size() As Byte: End Property
Public Property Get Instance() As IGeneric: Set Instance = Me: End Property
Public Property Get Comparer() As IGenericComparer: Set Comparer = IGenericComparer: End Property
Public Function Create(ByVal Value As Variant) As IGenericValue: Set Create = Me: End Function
Public Function CompareTo(ByVal Other As IGenericValue) As CompareResult
        
        If Other Is Nothing Then _
            CompareTo = CompareResult.IsGreater: Exit Function
        
        If Other Is Me Then _
            CompareTo = CompareResult.IsEqual: Exit Function

        CompareTo = CompareResult.IsLower
        
End Function
Public Function ToSqlParameter(Optional ByVal Direction As ISqlParameterDirection = ISqlParameterDirection.ParamInput, Optional ByVal Name As String) As Object: End Function

Private Sub Class_Initialize()
    Call System.Throw(Me, "Private Sub Class_Initialize()").Instantiation
End Sub

'IGenericCompare IGenericValue Variant comparer
Private Function IGenericComparer_GetHashCode(ByVal Obj As IGeneric) As Long: IGenericComparer_GetHashCode = Obj.HashValue: End Function
Private Function IGenericComparer_Compare(ByVal A As IGeneric, ByVal B As IGeneric) As CompareResult

    Dim x As IGenericValue, y As IGenericValue
    Dim i As Variant, j As Variant
    
    If A Is Nothing Then _
        IGenericComparer_Compare = IIf(B Is Nothing, CompareResult.IsEqual, CompareResult.IsLower): Exit Function
    
    If B Is Nothing Then _
        IGenericComparer_Compare = CompareResult.IsGreater: Exit Function
        
    Set x = A: i = x.ToValue: Set y = B: j = y.ToValue
    
    Select Case True
        Case i > j: IGenericValue_CompareTo = IsGreater
        Case i < j: IGenericValue_CompareTo = IsLower
        Case Else: IGenericValue_CompareTo = IsEqual
    End Select

End Function

Private Function IGenericComparer_Equals(ByVal A As IGeneric, ByVal B As IGeneric) As Boolean

    If x Is Nothing = False Then _
        IGenericComparer_Equals = x.Equals(y): Exit Function

    If y Is Nothing = False Then _
        IGenericComparer_Equals = False: Exit Function

    IGenericComparer_Equals = True: Exit Function

End Function

'IGeneric
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = True: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = IGenericValue: End Property
Private Property Get IGeneric_VirtualPointer() As LongPtr: End Property
Private Property Get IGeneric_VirtualSize() As Long: End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_HashValue() As Long: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = True: End Property
Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean: IGeneric_Equals = Other Is IGeneric: End Function
Private Function IGeneric_ToString() As String: IGeneric_ToString = "Class: " & TypeName$(IGenericValue) & " :: Default": End Function
Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = IGeneric: End Function
Private Sub IGeneric_Dispose(): End Sub

