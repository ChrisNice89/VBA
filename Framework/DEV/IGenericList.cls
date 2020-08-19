VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
Option Explicit
'@Folder "<T>Enumerable"

Public Property Get Item(ByVal Index As Long) As IGeneric: End Property
Attribute Item.VB_UserMemId = 0
Public Property Set Item(ByVal Index As Long, ByVal Object As IGeneric): End Property
Public Property Get Elements() As IGenericCollection: End Property

Public Function IndexOf(ByVal Item As IGeneric) As Long: End Function
Public Sub Insert(ByVal Index As Long, ByVal Item As IGeneric): End Sub
Public Sub RemoveAt(ByVal Index As Long): End Sub

'Private Sub Class_Initialize()
'    Call Skynet.Push(Me, "Private Sub Class_Initialize()").Instantiation.Raise
'End Sub
