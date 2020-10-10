VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISqlValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "<T>Class"

Public Enum AdoDataType
    T_String
    T_Numeric
    T_Float
    T_Date
    T_Boolean
End Enum

Public Function Convert(ByVal Value As Variant) As ISqlValue: End Function
Public Function FormatValue(ByVal Value As Variant) As String: End Function

Public Property Get Primitive() As Variant: End Property
Attribute Primitive.VB_UserMemId = 0
Public Property Get SqlType() As AdoDataType: End Property
Public Property Get PrimitiveType() As VbVarType: End Property
Public Property Get IsNumeric() As Boolean: End Property
Public Property Get Size() As Byte: End Property
Public Property Get Format() As String: End Property

'Private Sub Class_Initialize()
'    Call Skynet.Throw(Me, "Private Sub Class_Initialize()").Instantiation
'End Sub
