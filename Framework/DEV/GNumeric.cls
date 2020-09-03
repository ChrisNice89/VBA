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
Implements IPrototype
Implements IValue

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

Private Type Member
    Value As Currency
    Format As TNumericFormat
End Type
Private myClass As Member
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
    If myClass.Value > MaxLong Then
        ToLong = MaxLong
    Else
        ToLong = CLng(myClass.Value)
    End If
End Function

Public Function Minus(ByVal Other As GNumeric) As GNumeric: Set Minus = GNumeric.Build(myClass.Value - Other.Value): End Function
Public Function Plus(ByVal Other As GNumeric) As GNumeric: Set Plus = GNumeric.Build(myClass.Value + Other.Value): End Function
Public Function Modulo(ByVal m As Double) As GNumeric: Set Modulo = GNumeric((myClass.Value - (Int(myClass.Value / m) * m))): End Function
Public Function Add(ByVal Factor As Currency) As GNumeric: Set Add = GNumeric(myClass.Value + Factor): End Function

Public Property Get Format() As TNumericFormat: Format = myClass.Format: End Property
Public Property Get Value() As Currency: Value = myClass.Value: End Property
Public Property Get IsZero() As Boolean: IsZero = myClass.Value = 0: End Property
Public Property Get IsEven() As Boolean: IsEven = (myClass.Value - (Int(myClass.Value / 2) * 2)) = 0: End Property

Public Function Build(ByVal Value As Currency, Optional ByVal Format As TNumericFormat = TNumericFormat.DefaultNumber) As GNumeric
Attribute Build.VB_UserMemId = 0
    
    With myClass
        .Value = Value
        .Format = Format
    End With
    
    'Set Build = Skynet.Create(GNumeric)
    Set Build = Skynet.CreateInstance(New GNumeric, Virtual.Pointer, Virtual.Size)

End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Virtual.Size = LenB(myClass)
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
Private Property Get IValue_DatabaseType() As AdoDataType

End Property
Private Property Get IValue_Value() As Variant: IValue_Value = myClass.Value: End Property

Private Function IGeneric_Clone() As IGeneric
    
    Dim C As Member: C = myClass 'hardcopy

    Set IGeneric_Clone = Skynet.CreateInstance(New GNumeric, VarPtr(C), Virtual.Size)
   
End Function

Private Sub IGeneric_Dispose()
    myClass = EmptyStructure
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

    If Other Is Nothing Then IGeneric_Equals = False: Exit Function
    If Not Other.IsRelatedTo(Me) Then IGeneric_Equals = False: Exit Function

    Dim n As GNumeric: Set n = Other

    IGeneric_Equals = (myClass.Value = n.Value)
     
End Function

Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum
    
    If Other Is Nothing Then _
        IGeneric_CompareTo = IsGreater: Exit Function
    
    Dim n As GNumeric:  Set n = Other
    
    Select Case True
        Case myClass.Value > n.Value: IGeneric_CompareTo = IsGreater
        Case myClass.Value < n.Value: IGeneric_CompareTo = IsLower
        Case Else: IGeneric_CompareTo = IsEqual
    End Select

End Function

Private Property Get IGeneric_HashValue() As Long: IGeneric_HashValue = Skynet.LimitDouble(myClass.Value): End Property
Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean: IGeneric_IsRelatedTo = (TypeOf Other Is GNumeric): End Function
Private Function IGeneric_ToString() As String

    IGeneric_ToString = myClass.Value
    
End Function
    
Private Function IPrototype_CreateDefault() As IPrototype: Set IPrototype_CreateDefault = New GNumeric: End Function
Private Property Get IPrototype_StaticInstance() As Object: Set IPrototype_StaticInstance = GNumeric: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = Virtual.Size: End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property

