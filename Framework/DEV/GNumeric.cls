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
Implements IGenericValue
Implements IGenericIterator

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

'
Private Type Range
    Start As Long
    Last As Long
    i As Long
    current As IGeneric
End Type

Private Type Member
    Value As Currency
    Range As Range
End Type
Private myclass As Member
Private Formats() As String

Public Property Get min() As Currency: min = -922337203685477@: End Property
Public Property Get Max() As Currency: Max = MAXVALUE: End Property

Public Property Get Range(ByVal Start As Long, ByVal Last As Long) As IGenericIterator
    
    Dim c As Member
    
    If Start > Last Then _
        Call System.Throw(Me, "Public Property Get Range").InvalidInput("Start > Last")
        
    With c.Range
        .Start = Start
        .Last = Last
        .i = .Start - 1
    End With
   
    Set Range = System.CreateInstance(New GNumeric, VarPtr(c), LenB(c))

End Property

Public Property Get HasFraction() As Boolean: HasFraction = Not (Me.Modulus(myclass.Value, 1@) = 0): End Property
Public Property Get IsDefault() As Boolean: IsDefault = Me Is GNumeric: End Property
Public Property Get ToFormat(ByVal Format As NumericFormat) As String: ToFormat = myclass.Value: End Property
Public Property Get Value() As Currency: Value = myclass.Value: End Property
Public Property Get IsZero() As Boolean: IsZero = myclass.Value = 0: End Property
Public Property Get IsNegative() As Boolean: IsNegative = myclass.Value < 0: End Property
Public Property Get Fraction() As Currency: Fraction = myclass.Value - Int(myclass.Value): End Property
Public Property Get IsEven() As Boolean: IsEven = (Me.Modulus(myclass.Value, 2@) = 0): End Property

Public Property Get Default(Optional ByRef Value As Currency = DEFAULT_VALUE) As GNumeric

        If Not (Me Is GNumeric) Then _
            Set Default = GNumeric.Default(Value): Exit Property
        
        myclass.Value = Value
        Set Default = Me

End Property
Public Function Build(ByVal Value As Currency) As GNumeric
Attribute Build.VB_UserMemId = 0
    
    If Not (Me Is GNumeric) Then _
        Set Build = GNumeric.Build(Value): Exit Function
        
    myclass.Value = Value
    Set Build = New GNumeric: myclass = EmptyStructure

End Function

Private Sub Class_Initialize()
    With GNumeric
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
Public Function Floor() As GNumeric
    
    Set Floor = Me
    
    If Floor.HasFraction Then _
        Set Floor = GNumeric.Build(VBA.Int(myclass.Value))
  
End Function
'TODO Overflow
Public Function Ceiling() As GNumeric
    
    Set Ceiling = Me
    
    If Ceiling.HasFraction Then _
        Set Ceiling = GNumeric.Build(VBA.Int(myclass.Value) + 1) '-VBA.Int(-myclass.Value)
  
End Function

Public Function LOGn(ByVal Base As Byte) As Long: LOGn = Log(myclass.Value) / Log(Base): End Function
Public Function Multiply(ByVal Other As GNumeric) As GNumeric: Set Multiply = GNumeric.Build(myclass.Value * Other.Value): End Function
Public Function Minus(ByVal Other As GNumeric) As GNumeric: Set Minus = GNumeric.Build(myclass.Value - Other.Value): End Function
Public Function Plus(ByVal Other As GNumeric) As GNumeric: Set Plus = GNumeric.Build(myclass.Value + Other.Value): End Function
Public Function Modulus(ByRef A As Currency, ByRef B As Currency) As Currency: Modulus = (A - (Int(A / B) * B)): End Function
Public Function Modulo(ByVal Other As GNumeric) As GNumeric: Set Modulo = GNumeric.Build(Me.Modulus(myclass.Value, Other.Value)): End Function
Public Function Add(ByVal Factor As Currency) As GNumeric: Set Add = GNumeric.Build(myclass.Value + Factor): End Function
Public Function ToLong(ByVal V As Currency) As Long: ToLong = CLng(V - (Fix(V / MAXLONG) * MAXLONG)): End Function

'IGeneric
Private Property Get IGeneric_ClassName() As String: IGeneric_ClassName = TypeName$(Me): End Property
Private Property Get IGeneric_IsDefault() As Boolean: IGeneric_IsDefault = Me.IsDefault: End Property
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = GNumeric: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = myclass.Value: End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GNumeric):  End Property

Private Function EmptyStructure() As Member: End Function
Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = Me: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean
    
    If Other Is Nothing Then _
        IGeneric_Equals = False: Exit Function
    
    If Not TypeOf Other Is GNumeric Then _
        IGeneric_Equals = False: Exit Function
        
    Dim otherNumber As GNumeric
    Set otherNumber = Other

    IGeneric_Equals = (myclass.Value = otherNumber.Value)
    
End Function

Private Function IGeneric_ToString() As String: IGeneric_ToString = myclass.Value: End Function

'IGenericValue
Private Property Get IGenericValue_VarType() As VbVarType: IGenericValue_VarType = vbCurrency: End Property
Private Property Get IGenericValue_Size() As Byte: IGenericValue_Size = LenB(myclass.Value): End Property
Private Property Get IGenericValue_SqlType() As ISqlDataType: IGenericValue_SqlType = ISqlDataType.Numeric_Type: End Property
Private Property Get IGenericValue_ToValue() As Variant: IGenericValue_ToValue = myclass.Value: End Property
Private Property Get IGenericValue_Instance() As IGeneric: Set IGenericValue_Instance = Me: End Property
Private Property Get IGenericValue_Comparer() As IGenericComparer: Set IGenericValue_Comparer = IGenericComparer: End Property

Private Function IGenericValue_Convert(ByVal Value As Variant) As IGenericValue: Set IGenericValue_Convert = GNumeric.Build(VBA.IIf(VBA.Isnull(Value), DEFAULT_VALUE, Value)): End Function
Private Function IGenericValue_CompareTo(ByVal Other As IGenericValue) As CompareResult
    
    If Other Is Nothing Then _
        IGenericValue_CompareTo = IsGreater: Exit Function
    
    Dim N As GNumeric:  Set N = Other
    Dim i As Currency
    i = N.Value
    
    Select Case True
        Case myclass.Value > i: IGenericValue_CompareTo = IsGreater
        Case myclass.Value < i: IGenericValue_CompareTo = IsLower
        Case Else: IGenericValue_CompareTo = IsEqual
    End Select
    
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
        .Value = myclass.Value
    End With
    
    Set IGenericValue_ToSqlParameter = p
    
End Function

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Range.current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    With myclass.Range
        If .i = .Last Then _
            Set .current = Nothing: Set Out = Nothing: Exit Function
            
        Set .current = GNumeric(InterlockedIncrement(.i)): Set Out = .current
        IGenericIterator_HasNext = True
      
    End With
    
End Function

Private Sub IGenericIterator_Reset()
    Set myclass.Range.current = Nothing
    myclass.Range.i = myclass.Range.Start
End Sub

