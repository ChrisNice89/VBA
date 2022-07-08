VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed

'@IgnoreModule ImplementedInterfaceMember
'@PredeclaredId
'@Interface
'@Folder "Primitives.Abstract"
Option Explicit

'*======================================================================================================================
'*
'*          Copyright (c) Christoph Nitz.  All rights reserved.
'*
'*======================================================================================================================
'*
'* Class:  IGenericValue
'*
'* Purpose: Base Interface for all value typed objects
'*          This interface is implemented by types whose values can be ordered or sorted.
'*
'*
'*
'* Notes: All numeric types (such as Int and Numeric) implement IGenericValue, as do String, and Date. Custom types should also provide their own implementation of IGenericValue to enable object instances to be ordered or sorted.
'*        IGenericValues are immutable, which means we are creating new memory everytime instead of working on existing memory.
'*
'*
'*
'*======================================================================================================================

Implements IGeneric
Implements IGenericComparer
Implements IGenericHashProvider
Implements IGenericPredicator

Public Enum CompareResult
    IsLower = -1 'Less than zero -> Current instance precedes the object specified by the Compare method in the sort order.
    IsEqual = 0 ' Zero -> Current instance occurs in the same position in the sort order as the object specified by the Compare method.
    isgreater = 1 ' Greater than zero ->  Current instance follows the object specified by the Compare method in the sort order.
End Enum

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
    ParamInputOutput = 3 'Der Parameter ist sowohl Ãœbergabeparameter, liefert aber auch einen Wert zurück.
    ParamReturnValue = 4 'Beschreibt einen Return-Wert
End Enum

Public Property Get GetValueOrDefault(Optional ByVal DefaultValue As Variant = Null) As Variant: GetValueOrDefault = DefaultValue: End Property
'@DefaultMember
Public Property Get GetValue() As Variant: GetValue = Empty: End Property
Attribute GetValue.VB_UserMemId = 0
Public Property Get ToString() As String: ToString = "": End Property
Public Property Get HasValue() As Boolean: HasValue = False: End Property
Public Property Get SqlType() As ISqlDataType: SqlType = Variant_Type: End Property
Public Property Get VarType() As VBA.VbVarType: VarType = VBA.vbVariant: End Property
Public Property Get Size() As Byte: Size = 16: End Property

Public Function Of(ByRef Value As Variant) As IGenericValue: Set Of = Me: End Function

Public Function Equals(ByVal Other As IGenericValue) As Boolean
        
    If Other Is Nothing Then _
        Exit Function
   
    Equals = VBA.IIf(Other.HasValue, False, True)
    
End Function

Public Function HashCode() As Long: HashCode = 0: End Function
Public Function IsGreaterThan(ByVal Other As IGenericValue) As Boolean: End Function
Public Function IsLessThan(ByVal Other As IGenericValue) As Boolean: End Function
Public Function IsBetween(ByVal A As IGenericValue, ByVal B As IGenericValue) As Boolean: End Function

'TODO
'@Ignore FunctionReturnValueDiscarded
Public Function ToSqlParameter(Optional ByVal Direction As ISqlParameterDirection = ISqlParameterDirection.ParamInput, Optional ByVal Name As String) As Object: Call GenericError.Build(Me, "Public Function ToSqlParameter").NotImplemented: End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = True: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = Me: End Property
Private Property Get IGeneric_VirtualPointer() As LongPtr: IGeneric_VirtualPointer = 0: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = 0: End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_HashCode() As Long: IGeneric_HashCode = Me.HashCode: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean: IGeneric_IsRelatedTo = TypeOf Other Is IGenericValue: End Property
Private Function IGeneric_Equals(ByVal Other As Object) As Boolean: IGeneric_Equals = Me.Equals(Other): End Function
Private Function IGeneric_ToString() As String: IGeneric_ToString = Me.ToString: End Function

Private Function IGenericHashProvider_HashOf(ByVal Element As Object) As Long
    
    Dim Value As IGenericValue
    
    If Element Is Nothing Then _
        Exit Function
        
    Set Value = Element
    IGenericHashProvider_HashOf = Value.HashCode
    
End Function

Private Function IGenericHashProvider_Equality(ByVal A As Object, ByVal B As Object) As Boolean: IGenericHashProvider_Equality = IGenericComparer_Equality(A, B): End Function

Private Function IGenericPredicator_IsTrue(ByVal A As Object, ByVal B As Object) As Boolean: IGenericPredicator_IsTrue = IGenericComparer_AscendingOrdered(A, B): End Function

Private Function IGenericComparer_Equality(ByVal A As Object, ByVal B As Object) As Boolean

    Dim x As IGenericValue
    Dim y As IGenericValue

    If A Is Nothing Then _
        IGenericComparer_Equality = B Is Nothing: Exit Function
        
    If B Is Nothing Then _
        Exit Function
        
    Set x = A
    Set y = B
    
    IGenericComparer_Equality = (x.GetValue = y.GetValue)

End Function

Private Function IGenericComparer_AscendingOrdered(ByVal A As Object, ByVal B As Object) As Boolean
    
    Dim x As IGenericValue
    Dim y As IGenericValue
    
'    If A Is Nothing Then _
'       IGenericComparer_AscendingOrdered = Not B Is Nothing: Exit Function
'
'    If B Is Nothing Then _
'        Exit Function

    If A Is Nothing Then _
        Exit Function

    If B Is Nothing Then _
        IGenericComparer_AscendingOrdered = True: Exit Function
        
    Set x = A
    Set y = B
    
    IGenericComparer_AscendingOrdered = x.GetValue < y.GetValue
    
End Function
