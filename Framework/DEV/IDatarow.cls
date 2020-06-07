VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDatarow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Interface")
Option Compare Database

Public Property Get Field(ByVal Index As Variant) As IValue: End Property
Public Property Get Name(ByVal Index As Long) As String: End Property
Public Property Get Count() As Long: End Property

Public Function Update(ByVal Index As Variant, NewValue As IValue) As IDatarow: End Function
Public Sub Clear(): End Sub
Public Function Copy() As IDatarow: End Function

'Private Sub Class_Initialize()
'    Call Skynet.Push(Me, "Private Sub Class_Initialize()").Instantiation.Raise
'End Sub

