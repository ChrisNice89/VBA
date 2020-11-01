VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GNumeric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "<T>Class"
Option Explicit

Implements IGeneric

Private Const MaxLong As Long = 2 ^ 31 - 1

Public Enum TNumericFormat
    [_First]
    DefaultNumber
    TwoDecimals
    Thousands
    Fixed
    EuroCurrency
    DollarCurrency
    [_Last]
End Enum

'
Private Type Member
    Value As Currency
    Precision As Byte
    NumericScale As Byte
    Format As TNumericFormat
End Type
Private myclass As Member
Private Formats() As String

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Min() As Currency
    Min = -922337203685477#
End Property

Public Property Get Max() As Currency
    Max = 922337203685477#
End Property

Public Function ToLong() As Long
    If myclass.Value > MaxLong Then
        ToLong = MaxLong
    Else
        ToLong = CLng(myclass.Value)
    End If
End Function

Public Function Minus(ByVal Other As GNumeric) As GNumeric: Set Minus = GNumeric.Build(myclass.Value - Other.Value): End Function
Public Function Plus(ByVal Other As GNumeric) As GNumeric: Set Plus = GNumeric.Build(myclass.Value + Other.Value): End Function
Public Function Modulo(ByVal m As Double) As GNumeric: Set Modulo = GNumeric((myclass.Value - (Int(myclass.Value / m) * m))): End Function
Public Function Add(ByVal Factor As Currency) As GNumeric: Set Add = GNumeric(myclass.Value + Factor): End Function

Public Property Get Format() As TNumericFormat: Format = myclass.Format: End Property
Public Property Get Value() As Currency: Value = myclass.Value: End Property
Public Property Get IsZero() As Boolean: IsZero = myclass.Value = 0: End Property
Public Property Get IsEven() As Boolean: IsEven = (myclass.Value - (Int(myclass.Value / 2) * 2)) = 0: End Property

Public Function Build( _
                            ByVal Value As Currency, _
                            Optional ByVal Precision As Byte = 4, _
                            Optional ByVal NumericScale As Byte = 4, _
                            Optional ByVal Format As TNumericFormat = TNumericFormat.DefaultNumber) As GNumeric
Attribute Build.VB_UserMemId = 0
    With myclass
        .Value = Value
        .Precision = Precision
        .NumericScale = NumericScale
        .Format = Format
    End With
    
    'Set Build = Skynet.Create(GNumeric)
    Set Build = Skynet.CreateInstance(New GNumeric, Virtual.Pointer, Virtual.Size)

End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myclass)
    Virtual.Size = LenB(myclass)
'    ReDim Formats(TNumericFormat.[_First] To TNumericFormat.[_Last])
'    Formats(TNumericFormat.DefaultNumber) = "General Number"
'
' TNumericFormat.TwoDecimals VBA.Format(myClass.Value, "0.00"): Exit Function
'        Case TNumericFormat.Thousands
'            IObject_ToString = VBA.Format(myClass.Value, "#,##0.00"): Exit Function
'        Case TNumericFormat.Fixed
'            IObject_ToString = VBA.Format(myClass.Value, "#,##0.00"): Exit Function
'
''        ' cur1 is now equal to the String "$1,000,000.00".
''        cur2 = FormatCurrency(1000000, , , , vbFalse)
''        ' cur2 is now equal to the String "$1000000.00".
''        cur3 = FormatCurrency(100.55, 0)
''        ' cur3 is now equal to the String "$101".
''        cur4 = FormatCurrency(-500, 2, , vbTrue)
''        ' cur4 is now equal to the String "($500.00)".
    
End Sub

Private Property Get IValue_Object() As IGeneric: Set IValue_Object = Me: End Property
Private Property Get IValue_DatabaseType() As ISqlDataType

End Property
Private Property Get IValue_Value() As Variant: IValue_Value = myclass.Value: End Property

'IGeneric
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GNumeric: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = Virtual.Pointer: End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = Virtual.Size: End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = myclass.Value: End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GNumeric):  End Property

Private Sub IGeneric_Dispose(): End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Clone() As IGeneric
    
    Dim c As Member: c = myclass 'hardcopy

    Set IGeneric_Clone = Skynet.CreateInstance(New GNumeric, VarPtr(c), Virtual.Size)
   
End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

    If Other Is Nothing Then IGeneric_Equals = False: Exit Function

    Dim N As GNumeric: Set N = Other

    IGeneric_Equals = (myclass.Value = N.Value)
     
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    
    If Other Is Nothing Then _
        IGeneric_CompareTo = IsGreater: Exit Function
    
    Dim N As GNumeric:  Set N = Other
    Dim i As Currency
    i = N.Value
    
    Select Case True
        Case myclass.Value > i: IGeneric_CompareTo = IsGreater
        Case myclass.Value < i: IGeneric_CompareTo = IsLower
        Case Else: IGeneric_CompareTo = IsEqual
    End Select

End Function

Private Function IGeneric_ToString() As String

    IGeneric_ToString = myclass.Value
    
End Function
    
'IGenericValue
Private Function IGenericValue_ToSqlParameter(Optional ByVal Direction As ISqlParameterDirection = 1&, Optional ByVal Name As String) As Variant
    
    Dim P As ADODB.Parameter
    Set P = New ADODB.Parameter
    
    With P
        .Direction = Direction
        .Attributes = adParamNullable '64
        .Precision = myclass.Precision
        .NumericScale = myclass.NumericScale
        .Type = ISqlDataType.Decimal_Type
        .Name = Name
        .Value = myclass.Value
    End With
    
    Set IGenericValue_ToSqlParameter = P
    
End Function
