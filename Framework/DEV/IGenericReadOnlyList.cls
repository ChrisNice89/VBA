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


Public Property Get BaseIndex() As Long: End Property
Public Property Get Item(ByVal Index As Long) As IGeneric: End Property
Attribute Item.VB_UserMemId = 0
Public Property Get Count() As Long: End Property

Public Function IndexOf(ByVal Item As IGeneric) As Long: End Function
Public Function Iterator() As IGenericIterator: End Function

Public Sub CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): End Sub
Public Sub Clear(): End Sub

'Private Sub Class_Initialize()
'    Call Skynet.Throw(Me, "Private Sub Class_Initialize()").Instantiation
'End Sub
