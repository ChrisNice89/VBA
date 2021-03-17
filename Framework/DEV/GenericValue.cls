VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericValue"
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
'* Class:  GenericValue
'*
'* Purpose:
'*
'*
'*
'* Notes:
'*
'*
'*
'*
'*
'*======================================================================================================================

'@Folder "<T>Class"
Option Explicit
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

'Gibt an in welche Richtung der Parameter agiert. Die Konstanten stammen aus dem ParameterDirectoinEnum.
Public Enum ISqlParameterDirection
    ParamUnknown = 0 'Die Richtung des Parameters ist nicht bekannt.
    ParamInput = 1 'Es handelt sich um einen Input-Parameter
    ParamOutput = 2 'Es handelt sich um einen Output-Parameter
    ParamInputOutput = 3 'Der Parameter ist sowohl Übergabeparameter, liefert aber auch einen Wert zurück.
    ParamReturnValue = 4 'Beschreibt einen Return-Wert
End Enum

Private Type Member
    Value As Variant
End Type
Private myclass As Member

Public Property Get ToValue() As Variant: ToValue = myclass.Value: End Property
Attribute ToValue.VB_UserMemId = 0
Public Property Get SqlType() As ISqlDataType: SqlType = Variant_Type: End Property
Public Property Get VarType() As VbVarType: VarType = VBA.VarType(myclass.Value): End Property
Public Property Get Size() As Byte: Size = 16: End Property
Public Property Get Instance() As IGeneric: Set Instance = Me: End Property
Public Property Get Comparer() As IGenericComparer: Set Comparer = Me: End Property
Public Function Convert(ByVal Value As Variant) As GenericValue
    
    If (VBA.IsNull(Value)) Then _
        Set Convert = GenericValue: Exit Function
  
    myclass.Value = Value
    Set Convert = New GenericValue
    myclass.Value = Null
    
End Function
Private Function EmptyStructure() As Member: End Function

Public Function CompareTo(ByVal Other As GenericValue) As CompareResult
        
        Dim otherValue As GenericValue
        Dim i As Variant
        
        If Other Is Nothing Then _
            CompareTo = CompareResult.IsGreater: Exit Function
        
'        If Other Is Me Then _
'            CompareTo = CompareResult.IsEqual: Exit Function
        
        Set otherValue = Other
        i = otherValue.ToValue
    
        Select Case True
            Case i > myclass.Value: CompareTo = IsGreater
            Case i < myclass.Value: CompareTo = IsLower
            Case Else: CompareTo = IsEqual
        End Select
        
End Function

Public Function ToSqlParameter(Optional ByVal Direction As ISqlParameterDirection = ISqlParameterDirection.ParamInput, Optional ByVal Name As String) As Object
    Call System.Throw(Me, "Public Function ToSqlParameter").NotImplemented("Not implemented yet")
End Function

Private Sub Class_Initialize()
    myclass.Value = GenericValue.ToValue
End Sub

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me Is GenericValue: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GenericValue: End Property
Private Property Get IGeneric_VirtualPointer() As LongPtr: IGeneric_VirtualPointer = VarPtr(myclass.Value): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass.Value): End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_HashValue() As Long
    Call System.Throw(Me, "Private Property Get IGeneric_HashValue").NotImplemented("Not implemented yet")
End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = TypeOf Other Is GenericValue: End Property
Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

    Dim otherValue As GenericValue
        
    If Other Is Nothing Then _
        Exit Function
    
    If Not TypeOf Other Is GenericValue Then _
        Exit Function
        
    Set otherValue = Other
    
    IGeneric_Equals = myclass.Value = otherValue.ToValue
    
End Function
Private Function IGeneric_ToString() As String: IGeneric_ToString = IIf(VBA.IsNull(myclass.Value), "NULL", myclass.Value): End Function
Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = Me: End Function

'IGenericCompare Variant comparer
Private Function IGenericComparer_GetHashCode(ByVal Obj As IGeneric) As Long
    If Not (Obj Is Nothing) Then _
        IGenericComparer_GetHashCode = Obj.HashValue
End Function

Private Function IGenericComparer_Compare(ByVal A As IGeneric, ByVal B As IGeneric) As CompareResult

    Dim x As GenericValue, y As GenericValue
    Dim i As Variant, j As Variant
    
    If A Is B Then _
        IGenericComparer_Compare = CompareResult.IsEqual: Exit Function
        
    If A Is Nothing Then _
        IGenericComparer_Compare = IIf(B Is Nothing, CompareResult.IsEqual, CompareResult.IsLower): Exit Function
    
    If B Is Nothing Then _
        IGenericComparer_Compare = CompareResult.IsGreater: Exit Function
        
    Set x = A: i = x.ToValue: Set y = B: j = y.ToValue
    
    Select Case True
        Case i > j: IGenericComparer_Compare = IsGreater
        Case i < j: IGenericComparer_Compare = IsLower
        Case Else: IGenericComparer_Compare = IsEqual
    End Select

End Function

Private Function IGenericComparer_Equals(ByVal A As IGeneric, ByVal B As IGeneric) As Boolean
    
    Dim x As GenericValue, y As GenericValue
    
    If A Is B Then _
        IGenericComparer_Equals = True: Exit Function
        
    If A Is Nothing Or B Is Nothing Then _
        IGenericComparer_Equals = False: Exit Function
    
    Set x = A: Set y = B
    
    IGenericComparer_Equals = x.ToValue = y.ToValue

End Function
