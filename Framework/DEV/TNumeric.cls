VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TNumeric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "<T>Class"
Option Explicit

Implements IObject
Implements IPrototype
Implements IValue

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
    If myClass.Value > MAXLONG Then
        ToLong = MAXLONG
    Else
        ToLong = CLng(myClass.Value)
    End If
End Function

Public Function Minus(ByVal Other As TNumeric) As TNumeric
    Set Minus = TNumeric.Build(myClass.Value - Other.Value)
End Function

Public Function Plus(ByVal Other As TNumeric) As TNumeric
    Set Plus = TNumeric.Build(myClass.Value + Other.Value)
End Function

Public Function Add(ByVal Factor As Currency) As TNumeric
    Set Add = TNumeric.Build(myClass.Value + Factor)
End Function

Public Property Get Format() As TNumericFormat
    Format = myClass.Format
End Property

Public Property Get Value() As Currency
    Value = myClass.Value
End Property

Public Function Build(ByVal Value As Currency, Optional ByVal Format As TNumericFormat = TNumericFormat.DefaultNumber) As TNumeric
    With myClass
        .Value = Value
        .Format = Format
    End With
    
    Set Build = Skynet.Create(Me)
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    ReDim Formats(TNumericFormat.[_First] To TNumericFormat.[_Last])
    Formats(TNumericFormat.DefaultNumber) = "General Number"
      
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

Private Function IValue_Build(ByVal Value As Variant, Optional ByVal Format As Byte) As IObject
    
    With myClass
        .Value = Value
        .Format = Format
    End With
    
    Set IValue_Build = Skynet.Create(Me)
End Function
Private Property Get IValue_Object() As IObject: Set IValue_Object = Me: End Property
Private Property Get IValue_DatabaseType() As AdoDataType

End Property
Private Property Get IValue_Value() As Variant: IValue_Value = myClass.Value: End Property

Private Function IObject_Clone() As IObject
    
    Dim c As Member: c = myClass

    Virtual.Pointer = VarPtr(c)
    Set IObject_Clone = Skynet.Create(Me)
    Virtual.Pointer = VarPtr(myClass)
   
End Function
'TODO TMap IObject_Dispose
Private Sub IObject_Dispose()
    myClass = EmptyStructure
End Sub
Private Function EmptyStructure() As Member: End Function

Private Function IObject_Equals(Other As IObject) As Boolean

    If Other Is Nothing Then IObject_Equals = False: Exit Function
    If Not Other.IsRelatedTo(Me) Then IObject_Equals = False: Exit Function

    Dim n As TNumeric: Set n = Other

    IObject_Equals = (myClass.Value = n.Value)
     
End Function

'TODO TMap IObject_CompareTo
Private Function IObject_CompareTo(Other As IObject) As ComparerEnum
    Dim n As TNumeric
    
    If Other.IsRelatedTo(Me) Then
        Set n = Other
        Select Case True
            Case myClass.Value > n.Value: IObject_CompareTo = IsGreater '=  1     (String1 and String2 are not equal; String1 is greater than String2
            Case myClass.Value < n.Value: IObject_CompareTo = IsLower '= -1     (String1 and String2 are not equal; String1 is less than String2)
            Case Else: IObject_CompareTo = IsEqual '=  0     (String1 and String2 are equal)
        End Select
    End If
End Function

Private Property Get IObject_HashValue() As Long: IObject_HashValue = Skynet.LimitDouble(myClass.Value): End Property
Private Function IObject_IsRelatedTo(Other As IObject) As Boolean: IObject_IsRelatedTo = (TypeOf Other Is TString): End Function
Private Function IObject_ToString() As String: IObject_ToString = VBA.Format$(myClass.Value, Formats(Me.Format)): End Function
    
Private Property Get IPrototype_NewInstance() As IPrototype: Set IPrototype_NewInstance = New TNumeric: End Property
Private Property Get IPrototype_Size() As Long: IPrototype_Size = LenB(myClass): End Property
Private Property Get IPrototype_VirtualPointer() As LongPtr: IPrototype_VirtualPointer = Virtual.Pointer: End Property


