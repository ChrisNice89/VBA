VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gNumeric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@IgnoreModule UnassignedVariableUsage, VariableNotAssigned, SuspiciousPredeclaredInstanceAccess, ProcedureNotUsed, ConstantNotUsed, AssignedByValParameter
'@Folder "Primitives"
Option Explicit

Implements IGeneric
Implements IGenericValue

#If Win64 Then
    Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (lpAddend As Long) As Long
    Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (lpAddend As Long) As Long
#Else
    Private Declare Function InterlockedIncrement Lib "kernel32" (ByRef lpAddend As Long) As Long
    Private Declare Function InterlockedDecrement Lib "kernel32" (ByRef lpAddend As Long) As Long
#End If

Private Const MAX_LONG As Long = 2 ^ 31 - 1
Private Const MIN_VALUE As Currency = -922337203685477.5807@
Private Const MAX_VALUE As Currency = 922337203685477.5807@
Private Const NUMERIC_SCALE As Long = 8
Private Const PRECISION As Long = 4
Private Const DEFAULT_VALUE As Currency = 0@

'Public Enum NumericFormat
'    [_First]
'    DefaultNumber
'    TwoDecimals
'    Thousands
'    Fixed
'    EuroCurrency
'    DollarCurrency
'    [_Last]
'End Enum

Private Type Member
    Value As Currency
End Type
Private This As Member

Public Property Get Instance() As IGeneric: Set Instance = Me: End Property

Public Property Get MinValue() As Currency: MinValue = MIN_VALUE: End Property
Public Property Get MaxValue() As Currency: MaxValue = MAX_VALUE: End Property

Public Property Get HasFraction() As Boolean: HasFraction = Not (Me.Modulus(This.Value, 1@) = 0): End Property
Public Property Get IsDefault() As Boolean: IsDefault = Me Is gNumeric: End Property
Public Property Get IsZero() As Boolean: IsZero = This.Value = 0: End Property
Public Property Get IsNegative() As Boolean: IsNegative = Not ((This.Value And MIN_VALUE) = 0): End Property
Public Property Get Fraction() As Currency: Fraction = This.Value - VBA.Int(This.Value): End Property
Public Property Get IsEven() As Boolean: IsEven = (Me.Modulus(This.Value, 2@) = 0): End Property
Public Property Get Negate() As gNumeric: Set Negate = gNumeric.Of(-This.Value): End Property

'@DefaultMember
Public Property Get ToNumber() As Currency: ToNumber = This.Value: End Property
Attribute ToNumber.VB_UserMemId = 0
Public Function Of(ByVal Value As Currency) As gNumeric
    
    Static Zero As gNumeric
    
    If Not (Me Is gNumeric) Then _
        Set Of = gNumeric.Of(Value): Exit Function
    
    If Value = DEFAULT_VALUE Then
    
        If Zero Is Nothing Then
            This.Value = DEFAULT_VALUE
            Set Zero = New gNumeric
            This = EmptyStructure
        End If
    
        Set Of = Zero
        Exit Function
    End If
    
    This.Value = Value
    Set Of = New gNumeric
    This.Value = DEFAULT_VALUE
    
End Function

Private Sub Class_Initialize(): This.Value = gNumeric.ToNumber: End Sub
   
'    ReDim Formats(TNumericFormat.[_First] To TNumericFormat.[_Last])
'    Formats(TNumericFormat.DefaultNumber) = "General Number"
'
' TNumericFormat.TwoDecimals VBA.Format(This.Value, "0.00"): Exit Function
'        Case TNumericFormat.Thousands
'            IObject_ToString = VBA.Format(This.Value, "#,##0.00"): Exit Function
'        Case TNumericFormat.Fixed
'            IObject_ToString = VBA.Format(This.Value, "#,##0.00"): Exit Function
'
''        ' cur1 is now equal to the String "$1,000,000.00".
''        cur2 = FormatCurrency(1000000, , , , vbFalse)
''        ' cur2 is now equal to the String "$1000000.00".
''        cur3 = FormatCurrency(100.55, 0)
''        ' cur3 is now equal to the String "$101".
''        cur4 = FormatCurrency(-500, 2, , vbTrue)
''        ' cur4 is now equal to the String "($500.00)".

Public Function IsPrime() As Boolean
    
  Dim i As Currency
                    
  If This.Value <= 1 Then _
    Exit Function
  
  If This.Value = 2 Then _
    IsPrime = True: Exit Function
  
  If Me.IsEven Then _
    Exit Function
                                    
  If Me.HasFraction Then _
    Exit Function
  
  For i = 3 To VBA.Int(VBA.Sqr(This.Value)) Step 2
    If Me.Modulus(This.Value, i) = 0 Then _
      Exit Function
                                
  Next

  IsPrime = True
                                  
End Function
'TODO Overflow
'The floor() method rounds a number DOWNWARDS to the nearest integer
Public Function Floor() As gNumeric
    
    Set Floor = Me
    
    If Floor.HasFraction Then _
        Set Floor = gNumeric.Of(VBA.Int(This.Value))
  
End Function
'TODO Overflow
Public Function Ceiling() As gNumeric
    
    Set Ceiling = Me
    
    If Ceiling.HasFraction Then _
        Set Ceiling = gNumeric.Of(VBA.Int(This.Value) + 1) '-VBA.Int(-This.Value)
  
End Function

Public Function MaxOf(ByVal A As Currency, ByVal B As Currency) As gNumeric: Set MaxOf = gNumeric.Of(VBA.IIf(A > B, A, B)): End Function
Public Function MinOf(ByVal A As Currency, ByVal B As Currency) As gNumeric: Set MinOf = gNumeric.Of(VBA.IIf(A < B, A, B)): End Function
Public Function EqualsTo(ByVal Value As Currency) As Boolean: EqualsTo = This.Value = Value: End Function
Public Function CompareTo(ByVal Value As Currency) As CompareResult: CompareTo = Compare(This.Value, Value): End Function
Public Function Compare(ByVal A As Currency, ByVal B As Currency) As CompareResult: Compare = (A > B) - (A < B): End Function

Public Function IsBetween(ByVal A As Currency, B As Currency) As Boolean
    
    If (A > This.Value) Then _
        Exit Function
    
    If (B < This.Value) Then _
        Exit Function

    IsBetween = True
    
End Function

Public Function LOGn(ByVal Base As Byte) As gNumeric: Set LOGn = gNumeric.Of(VBA.Log(This.Value) / VBA.Log(Base)): End Function
Public Function Multiply(ByVal Value As Currency) As gNumeric: Set Multiply = gNumeric.Of(This.Value * Value): End Function
Public Function Minus(ByVal Value As Currency) As gNumeric: Set Minus = gNumeric.Of(This.Value - Value): End Function
Public Function Plus(ByVal Value As Currency) As gNumeric: Set Plus = gNumeric.Of(This.Value + Value): End Function
Public Function Modulus(ByRef A As Currency, ByRef B As Currency) As Currency:  Modulus = ((A - (VBA.Int(A / B) * B))): End Function
Public Function Modulo(ByVal Other As gNumeric) As gNumeric: Set Modulo = gNumeric.Of(Me.Modulus(This.Value, Other.ToNumber)): End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = VBA.TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me.IsDefault: End Property
Private Property Get IGeneric_GetType() As IGeneric: Set IGeneric_GetType = gNumeric: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(This): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(This): End Property
Private Property Get IGeneric_HashCode() As Long: IGeneric_HashCode = CLng(This.Value - (VBA.Fix(This.Value / MAX_LONG) * MAX_LONG)) Xor CLng((This.Value - VBA.Int(This.Value)) * 10000&): End Property

Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As Object) As Boolean

    If (Other Is Nothing) Then _
        Exit Property
        
    If (TypeOf Other Is gNumeric) Then _
        IGeneric_IsRelatedTo = True: Exit Property
        
    If (TypeOf Other Is IGenericValue) Then _
        IGeneric_IsRelatedTo = True: Exit Property

End Property

'@Ignore NonReturningFunction, EmptyMethod
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As Object) As Boolean
        
    If (Other Is Nothing) Then _
        Exit Function
        
    If TypeOf Other Is gNumeric Then
        Dim otherNumber As gNumeric: Set otherNumber = Other
        IGeneric_Equals = (This.Value = otherNumber.ToNumber)
    End If
    
End Function

Private Function IGeneric_ToString() As String: IGeneric_ToString = This.Value: End Function

'IGenericValue
Private Property Get IGenericValue_GetValueOrDefault(Optional ByVal DefaultValue As Variant = Null) As Variant: IGenericValue_GetValueOrDefault = VBA.IIf(Me.IsDefault, DefaultValue, This.Value): End Property
Private Property Get IGenericValue_HasValue() As Boolean: IGenericValue_HasValue = Not Me.IsDefault: End Property
Private Property Get IGenericValue_VarType() As VBA.VbVarType: IGenericValue_VarType = VBA.VbVarType.vbCurrency: End Property
Private Property Get IGenericValue_Size() As Byte: IGenericValue_Size = LenB(This.Value): End Property
Private Property Get IGenericValue_SqlType() As ISqlDataType: IGenericValue_SqlType = ISqlDataType.Numeric_Type: End Property
Private Property Get IGenericValue_GetValue() As Variant: IGenericValue_GetValue = This.Value: End Property
Private Property Get IGenericValue_ToString() As String: IGenericValue_ToString = This.Value: End Property

Private Function IGenericValue_HashCode() As Long: IGenericValue_HashCode = CLng(This.Value - (VBA.Fix(This.Value / MAX_LONG) * MAX_LONG)) Xor CLng((This.Value - VBA.Int(This.Value)) * 10000&): End Function
Private Function IGenericValue_Of(ByRef Value As Variant) As IGenericValue: Set IGenericValue_Of = gNumeric.Of(Value): End Function

Private Function IGenericValue_Equals(ByVal Other As IGenericValue) As Boolean: IGenericValue_Equals = (This.Value = Other.GetValue): End Function
Private Function IGenericValue_IsLessThan(ByVal Other As IGenericValue) As Boolean: IGenericValue_IsLessThan = (This.Value < Other.GetValue): End Function
Private Function IGenericValue_IsGreaterThan(ByVal Other As IGenericValue) As Boolean: IGenericValue_IsGreaterThan = (This.Value > Other.GetValue): End Function

Private Function IGenericValue_IsBetween(ByVal A As IGenericValue, ByVal B As IGenericValue) As Boolean
    
    If (A Is Nothing) Then _
        Exit Function
    
    If (B Is Nothing) Then _
        Exit Function
    
    If (A.HasValue) Then _
        If (B.HasValue) Then _
            IGenericValue_IsBetween = Me.IsBetween(A.GetValue, B.GetValue)
    
End Function

Private Function IGenericValue_ToSqlParameter(Optional ByVal Direction As ISqlParameterDirection = 1&, Optional ByVal Name As String) As Object
    
    Dim p As ADODB.Parameter
    Set p = New ADODB.Parameter
    
    With p
        .Direction = Direction
        .Attributes = adParamNullable '64
        .PRECISION = PRECISION
        .NumericScale = NUMERIC_SCALE
        .Type = IGenericValue_VarType
        .Name = Name
        .Value = This.Value
    End With
    
    Set IGenericValue_ToSqlParameter = p
    
End Function
