VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Gnumeric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Implements IGeneric
Implements GenericValue

#If Win64 Then
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

Private Const MAXVALUE As Currency = 922337203685477@
Private Const MAXLONG As Long = 2 ^ 31 - 1
Private Const NUMERIC_SCALE As Long = 8
Private Const PRECISION As Long = 4
Private Const DEFAULT_VALUE As Currency = 0@

Public Enum NumericFormat
    [_First]
    DefaultNumber
    TwoDecimals
    Thousands
    Fixed
    EuroCurrency
    DollarCurrency
    [_Last]
End Enum

Private Type Member
    Value As Currency
End Type
Private myclass As Member
Private Formats() As String

Public Property Get Min() As Currency: Min = -922337203685477@: End Property
Public Property Get Max() As Currency: Max = MAXVALUE: End Property

Public Property Get HasFraction() As Boolean: HasFraction = Not (Me.Modulus(myclass.Value, 1@) = 0): End Property
Public Property Get IsDefault() As Boolean: IsDefault = Me Is Gnumeric: End Property
Public Property Get ToFormat(ByVal Format As NumericFormat) As String: ToFormat = myclass.Value: End Property
Public Property Get Value() As Currency: Value = myclass.Value: End Property
Public Property Get IsZero() As Boolean: IsZero = myclass.Value = 0: End Property
Public Property Get IsNegative() As Boolean: IsNegative = myclass.Value < 0: End Property
Public Property Get Fraction() As Currency: Fraction = myclass.Value - Int(myclass.Value): End Property
Public Property Get IsEven() As Boolean: IsEven = (Me.Modulus(myclass.Value, 2@) = 0): End Property

Public Function Of(ByVal Value As Currency) As Gnumeric
Attribute Of.VB_UserMemId = 0
    
    Static Zero As Gnumeric
    
    If Not (Me Is Gnumeric) Then _
        Set Of = Gnumeric.Of(Value): Exit Function
    
    If Value = DEFAULT_VALUE Then
    
        If Zero Is Nothing Then
            myclass.Value = DEFAULT_VALUE
            Set Zero = New Gnumeric
            myclass = EmptyStructure
        End If
    
        Set Of = Zero
        Exit Function
    End If
    
    myclass.Value = Value
    Set Of = New Gnumeric
    myclass = EmptyStructure

End Function

Private Sub Class_Initialize()
    With Gnumeric
        myclass.Value = .Value
    End With
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

Public Function IsPrime() As Boolean
    
  Dim i As Currency
                    
  If myclass.Value <= 1 Then _
    Exit Function
  
  If myclass.Value = 2 Then _
    IsPrime = True: Exit Function
  
  If Me.IsEven Then _
    Exit Function
                                    
  If Me.HasFraction Then _
    Exit Function
  
  For i = 3 To VBA.Int(VBA.sqr(myclass.Value)) Step 2
    If Me.Modulus(myclass.Value, i) = 0 Then _
      Exit Function
                                
  Next

  IsPrime = True
                                  
End Function
'TODO Overflow
'The floor() method rounds a number DOWNWARDS to the nearest integer
Public Function Floor() As Gnumeric
    
    Set Floor = Me
    
    If Floor.HasFraction Then _
        Set Floor = Gnumeric.Of(VBA.Int(myclass.Value))
  
End Function
'TODO Overflow
Public Function Ceiling() As Gnumeric
    
    Set Ceiling = Me
    
    If Ceiling.HasFraction Then _
        Set Ceiling = Gnumeric.Of(VBA.Int(myclass.Value) + 1) '-VBA.Int(-myclass.Value)
  
End Function

Public Function LOGn(ByVal Base As Byte) As Long: LOGn = Log(myclass.Value) / Log(Base): End Function
Public Function Multiply(ByVal Other As Gnumeric) As Gnumeric: Set Multiply = Gnumeric.Of(myclass.Value * Other.Value): End Function
Public Function Minus(ByVal Other As Gnumeric) As Gnumeric: Set Minus = Gnumeric.Of(myclass.Value - Other.Value): End Function
Public Function Plus(ByVal Other As Gnumeric) As Gnumeric: Set Plus = Gnumeric.Of(myclass.Value + Other.Value): End Function
Public Function Modulus(ByRef A As Currency, ByRef B As Currency) As Currency: Modulus = (A - (Int(A / B) * B)): End Function
Public Function Modulo(ByVal Other As Gnumeric) As Gnumeric: Set Modulo = Gnumeric.Of(Me.Modulus(myclass.Value, Other.Value)): End Function
Public Function Add(ByVal Factor As Currency) As Gnumeric: Set Add = Gnumeric.Of(myclass.Value + Factor): End Function
Public Function ToLong(ByVal V As Currency) As Long: ToLong = CLng(V - (Fix(V / MAXLONG) * MAXLONG)): End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me.IsDefault: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = Gnumeric: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = CLng(myclass.Value - (Fix(myclass.Value / MAXLONG) * MAXLONG)): End Property

Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is Gnumeric):  End Property

Private Function EmptyStructure() As Member: End Function
Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = Me: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    
    If Other Is Nothing Then _
        IGeneric_Equals = False: Exit Function
    
    If Not TypeOf Other Is Gnumeric Then _
        IGeneric_Equals = False: Exit Function
        
    Dim otherNumber As Gnumeric
    Set otherNumber = Other

    IGeneric_Equals = (myclass.Value = otherNumber.Value)
    
End Function

Private Function IGeneric_ToString() As String: IGeneric_ToString = myclass.Value: End Function

'GenericValue
Private Property Get GenericValue_VarType() As VbVarType: GenericValue_VarType = vbCurrency: End Property
Private Property Get GenericValue_Size() As Byte: GenericValue_Size = LenB(myclass.Value): End Property
Private Property Get GenericValue_SqlType() As ISqlDataType: GenericValue_SqlType = ISqlDataType.Numeric_Type: End Property
Private Property Get GenericValue_ToValue() As Variant: GenericValue_ToValue = myclass.Value: End Property
Private Property Get GenericValue_Instance() As IGeneric: Set GenericValue_Instance = Me: End Property
Private Property Get GenericValue_Comparer() As IGenericComparer: Set GenericValue_Comparer = IGenericComparer: End Property

Private Function GenericValue_Convert(ByRef Value As Variant) As GenericValue

    If VBA.IsNull(Value) Then _
        Set GenericValue_Convert = IGeneric_Default: Exit Function
        
    Set GenericValue_Convert = Gnumeric.Of(CCur(Value))
    
End Function

Private Function GenericValue_CompareTo(ByVal Other As GenericValue) As CompareResult
    
    If Other Is Nothing Then _
        GenericValue_CompareTo = isgreater: Exit Function
    
    Dim otherNumber As Gnumeric:  Set otherNumber = Other
    Dim i As Currency: i = otherNumber.Value
    
    Select Case True
        Case myclass.Value > i: GenericValue_CompareTo = isgreater
        Case myclass.Value < i: GenericValue_CompareTo = islower
        Case Else: GenericValue_CompareTo = IsEqual
    End Select
    
End Function

Private Function GenericValue_ToSqlParameter(Optional ByVal Direction As ISqlParameterDirection = 1&, Optional ByVal Name As String) As Object
    
    Dim P As ADODB.Parameter
    Set P = New ADODB.Parameter
    
    With P
        .Direction = Direction
        .Attributes = adParamNullable '64
        .PRECISION = PRECISION
        .NumericScale = NUMERIC_SCALE
        .Type = GenericValue_VarType
        .Name = Name
        .Value = myclass.Value
    End With
    
    Set GenericValue_ToSqlParameter = P
    
End Function
