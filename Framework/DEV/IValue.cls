VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Interface")
Option Compare Database

Public Enum AdoDataType
    T_String
    T_Numeric
    T_Float
    T_Date
    T_Boolean
End Enum

Public Function Build(ByVal Value As Variant, Optional ByVal Format As Byte) As IObject: End Function
Public Property Get Value() As Variant: End Property
Public Property Get DatabaseType() As AdoDataType: End Property
Public Property Get Object() As IObject: End Property

'Private Sub Class_Initialize()
'    Call Skynet.Push(Me, "Private Sub Class_Initialize()").Instantiation.Raise
'End Sub

