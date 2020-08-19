VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
Option Explicit
'@Folder "<T>Enumerable"

Public Property Get IsFixedSize() As Boolean: End Property
Public Property Get IsReadOnly() As Boolean: End Property
Public Property Get Count() As Long: End Property

Public Function NewEnumerator() As IGenericEnumerator: End Function
Public Function ToArray() As GenericArray: End Function
Public Function Contains(ByVal Item As IGeneric) As Boolean: End Function
Public Sub CopyTo(ByVal GArray As GenericArray, ByVal ArrayIndex As Long): End Sub
Public Sub Clear(): End Sub

'Private Sub Class_Initialize()
'    Call Skynet.Push(Me, "Private Sub Class_Initialize()").Instantiation.Raise
'End Sub
