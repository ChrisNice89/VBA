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

Private Const MAXVALUE As Double = 922337203685477#
Private Const MAXLONG As Long = 2 ^ 31 - 1

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
Private Type Range
    Start As Long
    Last As Long
    i As Long
    Current As IGeneric
End Type

Private Type Member
    Value As Currency
    Precision As Byte
    NumericScale As Byte
    Range As Range
    Format As TNumericFormat
End Type
Private myclass As Member
Private Formats() As String

Public Property Get Min() As Currency
    Min = -922337203685477#
End Property

Public Property Get Max() As Currency
    Max = MAXVALUE
End Property

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

Public Property Get Format() As TNumericFormat: Format = myclass.Format: End Property
Public Property Get Value() As Currency: Value = myclass.Value: End Property
Public Property Get IsZero() As Boolean: IsZero = myclass.Value = 0: End Property
Public Property Get Precision() As Long: Precision = myclass.Precision: End Property
Public Property Get NumericScale() As Long: NumericScale = myclass.NumericScale: End Property
Public Property Get IsEven() As Boolean: IsEven = (myclass.Value - (Int(myclass.Value / 2) * 2)) = 0: End Property

Public Function Build( _
                            ByVal Value As Variant, _
                            Optional ByVal Precision As Byte = 4, _
                            Optional ByVal NumericScale As Byte = 4, _
                            Optional ByVal Format As TNumericFormat = TNumericFormat.DefaultNumber) As GNumeric
Attribute Build.VB_UserMemId = 0
    With myclass
        If Not IsNull(Value) Then _
            .Value = Value
         
        .Precision = Precision
        .NumericScale = NumericScale
        .Format = Format
    End With
    
    Set Build = New GNumeric: myclass = EmptyStructure

End Function

Private Sub Class_Initialize()
    With GNumeric
        myclass.Value = .Value
        myclass.Format = .Format
        myclass.Precision = .Precision
        myclass.NumericScale = .NumericScale
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

Public Function LOGn(ByVal Value As Long, Optional ByVal Base As Byte = 2) As Long: LOGn = Log(Value) / Log(Base): End Function
Public Function Multiply(ByVal Other As GNumeric) As GNumeric: Set Multiply = GNumeric.Build(myclass.Value * Other.Value): End Function
Public Function Minus(ByVal Other As GNumeric) As GNumeric: Set Minus = GNumeric.Build(myclass.Value - Other.Value): End Function
Public Function Plus(ByVal Other As GNumeric) As GNumeric: Set Plus = GNumeric.Build(myclass.Value + Other.Value): End Function
Public Function Modulo(ByVal m As Double) As GNumeric: Set Modulo = GNumeric((myclass.Value - (Int(myclass.Value / m) * m))): End Function
Public Function Add(ByVal Factor As Currency) As GNumeric: Set Add = GNumeric(myclass.Value + Factor): End Function
Public Function ToLong(ByVal V As Currency) As Long: ToLong = CLng(V - (Fix(V / MAXLONG) * MAXLONG)): End Function

'IGeneric
Private Property Get IGeneric_Default() As IGeneric: Set IGeneric_Default = New GNumeric: End Property
Private Property Get IGeneric_VirtualPointer() As Long: IGeneric_VirtualPointer = VarPtr(myclass): End Property
Private Property Get IGeneric_VirtualSize() As Long: IGeneric_VirtualSize = LenB(myclass): End Property
Private Property Get IGeneric_HashValue() As Long:  IGeneric_HashValue = myclass.Value: End Property
Private Property Get IGeneric_IsMutable() As Boolean: IGeneric_IsMutable = False: End Property
Private Property Get IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GNumeric):  End Property

Private Sub IGeneric_Dispose(): End Sub
Private Function EmptyStructure() As Member: End Function
Private Function IGeneric_Clone() As IGeneric: Set IGeneric_Clone = GNumeric.Build(myclass.Value, myclass.Precision, myclass.NumericScale, myclass.Format): End Function
Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean: IGeneric_Equals = IGenericValue_Equals(Other): End Function
Private Function IGeneric_ToString() As String: IGeneric_ToString = myclass.Value: End Function

'IGenericValue
Private Property Get IGenericValue_VarType() As VbVarType: IGenericValue_VarType = vbCurrency: End Property
Private Property Get IGenericValue_Size() As Byte: IGenericValue_Size = LenB(myclass.Value): End Property
Private Property Get IGenericValue_SqlType() As ISqlDataType: IGenericValue_SqlType = ISqlDataType.Decimal_Type: End Property
Private Property Get IGenericValue_ToValue() As Variant: IGenericValue_ToValue = myclass.Value: End Property
Private Property Get IGenericValue_Instance() As IGeneric: Set IGenericValue_Instance = Me: End Property
Private Property Get IGenericValue_Comparer() As IGenericComparer: Set IGenericValue_Comparer = IGenericValue.Comparer: End Property

Private Function IGenericValue_Create(ByVal Value As Variant, Optional ByVal Direction As ISqlParameterDirection = 1&) As IGenericValue
    Set IGenericValue_Create = Me.Build(Value)
End Function

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

Private Function IGenericValue_Equals(ByVal Other As IGenericValue) As Boolean
    
    If Other Is Nothing Then IGenericValue_Equals = False: Exit Function
    IGenericValue_Equals = (myclass.Value = Other.ToValue)
    
End Function

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

'IGenericIterator
Private Property Get IGenericIterator_Current() As IGeneric: Set IGenericIterator_Current = myclass.Range.Current: End Property
Private Function IGenericIterator_HasNext(ByRef Out As IGeneric) As Boolean
    
    With myclass.Range
        If .i = .Last Then _
            Set .Current = Nothing: Set Out = Nothing: Exit Function
            
        Set .Current = GNumeric(InterlockedIncrement(.i)): Set Out = .Current
        IGenericIterator_HasNext = True
      
    End With
    
End Function

Private Sub IGenericIterator_Reset()
    Set myclass.Range.Current = Nothing
    myclass.Range.i = myclass.Range.Start
End Sub

