VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericReadOnlyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("<T>Enumerable")
Option Compare Database

Public Property Get Item(ByVal Index As Long) As IGeneric: End Property
Attribute Item.VB_UserMemId = 0
Public Property Get Elements() As IGenericCollection: End Property

Public Function IndexOf(ByVal Item As IGeneric) As Long: End Function

'Private Sub Class_Initialize()
'    Call Skynet.Push(Me, "Private Sub Class_Initialize()").Instantiation.Raise
'End Sub
